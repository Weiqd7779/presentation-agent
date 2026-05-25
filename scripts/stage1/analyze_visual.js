const fs = require('fs');
const path = require('path');
const { callVisionLLM } = require('../../vision_llm_client.js');

function readJson(filePath, fallback) {
  if (!fs.existsSync(filePath)) return fallback;
  return JSON.parse(fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, ''));
}

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(`Vision LLM timed out after ${ms}ms`)), ms)),
  ]);
}

function shapeContext(slide) {
  if (!slide || !slide.shapes || !slide.shapes.length) return '- none';
  return slide.shapes.map((shape, index) => [
    `${index + 1}. ${shape.name}`,
    `   type: ${shape.type}`,
    `   text: ${shape.text || '(none)'}`,
    `   region: ${shape.region}`,
    `   geometry_in: x=${shape.geometry_in.x}, y=${shape.geometry_in.y}, w=${shape.geometry_in.w}, h=${shape.geometry_in.h}`,
    `   deterministic_role_hint: ${shape.likely_role}`,
  ].join('\n')).join('\n');
}

function fallbackSlideSpec(slide) {
  const lines = [];
  lines.push(`### Slide ${slide.slide}: Vision analysis unavailable`);
  lines.push('');
  lines.push('- Layout: Use `placeholder_layout.json` and the annotated slide image for placement.');
  lines.push('- Editable shapes:');
  for (const shape of slide.shapes || []) {
    lines.push(`  - \`${shape.name}\` (${shape.type}, ${shape.region}, ${shape.geometry_in.x}, ${shape.geometry_in.y}, ${shape.geometry_in.w}, ${shape.geometry_in.h} in): likely role \`${shape.likely_role}\`${shape.text ? `; current text \`${shape.text}\`` : ''}`);
  }
  lines.push('- Mapping notes: Shape names are exact. For generic Canva names, use the annotated image number and geometry to identify the visual role.');
  lines.push('- Risks: Vision role mapping did not complete for this slide; verify content placement during Stage 3 QA.');
  lines.push('');
  return lines.join('\n');
}

async function analyzeSlide(templateName, slide, timeoutMs) {
  const prompt = `
You are analyzing one rendered PowerPoint template slide for a content injection pipeline.
The image is annotated with bounding boxes and shape names. Canva and exported templates often use generic names, so infer each shape's visual role from its position and the annotated image.

Return Markdown only using this exact structure:

### Slide ${slide.slide}: short purpose
- Layout: one concise sentence.
- Editable shapes:
  - \`Exact shape name\`: role, position, recommended content, cautions.
- Do not edit:
  - shape names or visual areas that look decorative.
- Mapping notes: concrete Stage 2 mapping guidance.
- Risks: likely overflow, dense layout, confusing duplicate placeholders, or image crop risks.

Template: ${templateName}
Slide: ${slide.slide}
Deterministic XML geometry context:

${shapeContext(slide)}
`;

  const imagePath = slide.annotated_image || slide.image;
  if (!imagePath || !fs.existsSync(imagePath)) {
    throw new Error(`Annotated slide image not found for slide ${slide.slide}`);
  }
  return await withTimeout(callVisionLLM(prompt, [imagePath]), timeoutMs);
}

function buildRoleEntry(slide, shape, status, specPath) {
  return {
    type: shape.type,
    text: shape.text || '',
    region: shape.region,
    geometry_in: shape.geometry_in,
    deterministic_role_hint: shape.likely_role,
    vision_status: status,
    slide_spec: specPath ? path.basename(specPath) : '',
  };
}

async function main() {
  if (process.argv.length < 5) {
    console.error('Usage: node analyze_visual.js <template_file> <output_dir> <placeholders_file>');
    process.exit(1);
  }

  const templateFile = process.argv[2];
  const outputDir = process.argv[3];
  const templateName = path.basename(templateFile, path.extname(templateFile));
  const layoutPath = path.join(outputDir, 'placeholder_layout.json');
  const layout = readJson(layoutPath, { slides: [] });
  const timeoutMs = Number(process.env.PRESENTATION_AGENT_VISION_TIMEOUT_MS || 300000);
  const skipVision = String(process.env.PRESENTATION_AGENT_SKIP_VISION || '').toLowerCase() === '1';
  const specSlideDir = path.join(outputDir, 'spec_slides');
  fs.mkdirSync(specSlideDir, { recursive: true });

  const roles = {
    template: templateName,
    analysis_mode: 'annotated per-slide images with deterministic XML geometry fallback',
    timeout_ms_per_slide: timeoutMs,
    slides: {},
  };

  const lines = [];
  lines.push(`# Presentation Analysis Specification: ${templateName}`);
  lines.push('');
  lines.push('## Overall Structure');
  lines.push('');
  lines.push(`- Total slides: ${layout.slides.length || 'unknown'}`);
  lines.push('- Analysis mode: annotated per-slide images plus XML geometry');
  lines.push(`- Vision timeout per slide: ${timeoutMs} ms`);
  lines.push(`- Vision skipped: ${skipVision ? 'yes' : 'no'}`);
  lines.push('- Supporting files: `placeholder_layout.json`, `placeholder_roles.json`, `annotated_slides/`, `spec_slides/`');
  lines.push('');
  lines.push('## Slide Analysis');
  lines.push('');

  for (const slide of layout.slides) {
    const slideSpecPath = path.join(specSlideDir, `slide_${String(slide.slide).padStart(3, '0')}.md`);
    let status = 'fallback';
    let slideSpec = '';

    if (!skipVision && fs.existsSync(slideSpecPath)) {
      slideSpec = fs.readFileSync(slideSpecPath, 'utf8');
      status = 'cached';
    } else if (!skipVision) {
      try {
        slideSpec = (await analyzeSlide(templateName, slide, timeoutMs)).trim() + '\n';
        fs.writeFileSync(slideSpecPath, slideSpec, 'utf8');
        status = 'vision_ok';
      } catch (error) {
        console.error(`Warning: Vision LLM analysis failed for slide ${slide.slide}. ${error.message || error}`);
        slideSpec = fallbackSlideSpec(slide);
        fs.writeFileSync(slideSpecPath, slideSpec, 'utf8');
      }
    } else {
      slideSpec = fallbackSlideSpec(slide);
      fs.writeFileSync(slideSpecPath, slideSpec, 'utf8');
      status = 'skipped';
    }

    roles.slides[String(slide.slide)] = {};
    for (const shape of slide.shapes || []) {
      roles.slides[String(slide.slide)][shape.name] = buildRoleEntry(slide, shape, status, slideSpecPath);
    }

    lines.push(slideSpec.trim());
    lines.push('');
  }

  lines.push('## Mapping Guidance');
  lines.push('');
  lines.push('- Treat `placeholders.txt` as authoritative for exact text box and picture names.');
  lines.push('- Use `placeholder_layout.json` and annotated slide images to identify generic Canva names by position.');
  lines.push('- Use `placeholder_roles.json` as the machine-readable role map for Stage 2 planning.');
  lines.push('- If a slide falls back because Vision timed out, its XML geometry and annotated image are still usable for mapping.');
  lines.push('- Keep geometry in the template; Stage 2 should fill content, reorder/duplicate slides, and replace existing pictures.');

  fs.writeFileSync(path.join(outputDir, 'placeholder_roles.json'), JSON.stringify(roles, null, 2), 'utf8');
  console.log(lines.join('\n'));
}

main();
