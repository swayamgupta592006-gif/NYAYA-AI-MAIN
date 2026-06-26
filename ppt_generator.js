// ═══════════════════════════════════════════════════════════════
// NyayaMitra PPT GENERATOR — Government-Accepted Format
// Uses PptxGenJS to create official Indian legal presentation decks
// ═══════════════════════════════════════════════════════════════

// Kanon API key integration
const KANON_API_KEY = 'f669d3757ad0e8311628136459e86b3f58e51244';
const GEMINI_KEY_PRIMARY = 'AIzaSyD9ZDfi7uObfHHdjCXdXh_V2H1PJqRb6sU';

// ── MULTI-AI CALL (Kanon + Gemini fallback) ───────────────────
async function callKanonAI(prompt, systemPrompt = '') {
  // Try Kanon API first (Gemini via Kanon proxy)
  try {
    const response = await fetch('https://api.kanonaigc.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${KANON_API_KEY}`
      },
      body: JSON.stringify({
        model: 'gemini-2.0-flash',
        messages: [
          ...(systemPrompt ? [{ role: 'system', content: systemPrompt }] : []),
          { role: 'user', content: prompt }
        ],
        temperature: 0.2,
        max_tokens: 4000
      })
    });
    if (response.ok) {
      const data = await response.json();
      const text = data?.choices?.[0]?.message?.content;
      if (text) { console.log('✅ Kanon AI responded'); return text; }
    }
  } catch (e) {
    console.warn('Kanon API failed, falling back to Gemini:', e.message);
  }

  // Fallback: direct Gemini
  return await callGeminiDirect(prompt);
}

async function callGeminiDirect(prompt) {
  const MODELS = ['gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-2.5-flash'];
  for (const model of MODELS) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_KEY_PRIMARY}`;
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ role: 'user', parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.2 }
        })
      });
      const data = await res.json();
      if (!res.ok) continue;
      const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (text) return text;
    } catch (e) { continue; }
  }
  throw new Error('All AI models exhausted');
}

// ── PPT SLIDE DEFINITIONS (Government Standard Format) ───────────
// Indian government presentations follow NIC (National Informatics Centre) guidelines:
// - A4 / Widescreen 16:9
// - Official header with emblem position
// - Specific color: Navy Blue (#003087), Saffron (#FF6600), White background
// - Helvetica / Arial fonts
// - Numbered slides with department/ministry watermark

const GOV_COLORS = {
  navyBlue: '00308F',
  saffron:  'FF6600',
  gold:     'B8960C',
  white:    'FFFFFF',
  lightGray:'F5F5F5',
  darkGray: '333333',
  midGray:  '666666',
  ashoka:   '0047AB',
  green:    '138808',
  accent:   '003087',
};

// Government document type to PPT slide mapping
const DOC_PPT_TEMPLATES = {
  'FIR Draft Generator': {
    title: 'FIRST INFORMATION REPORT',
    subtitle: 'Under Section 154 CrPC | Indian Police Service',
    icon: '🚔',
    color: GOV_COLORS.navyBlue,
    slides: ['Cover', 'Complainant Details', 'Incident Summary', 'Accused Details', 'Legal Sections', 'Relief Sought', 'Signatures & Verification']
  },
  'RTI Application': {
    title: 'RIGHT TO INFORMATION APPLICATION',
    subtitle: 'Under RTI Act, 2005 | Ministry of Law and Justice',
    icon: '📋',
    color: GOV_COLORS.ashoka,
    slides: ['Cover', 'Applicant Details', 'Department Addressed', 'Information Sought', 'Grounds', 'Enclosures', 'Submission Details']
  },
  'Legal Notice': {
    title: 'LEGAL NOTICE',
    subtitle: 'Under Indian Contract Act / Civil Procedure Code',
    icon: '⚖️',
    color: GOV_COLORS.navyBlue,
    slides: ['Cover', 'Sender Details', 'Recipient Details', 'Subject Matter', 'Legal Grounds', 'Demand / Relief', 'Response Deadline']
  },
  'Rent Agreement': {
    title: 'RENT / LEAVE & LICENCE AGREEMENT',
    subtitle: 'Under Transfer of Property Act, 1882 / Registration Act',
    icon: '🏠',
    color: GOV_COLORS.green,
    slides: ['Cover', 'Landlord Details', 'Tenant Details', 'Property Description', 'Rent & Deposit Terms', 'Conditions & Covenants', 'Signature & Stamp']
  },
  'Consumer Complaint': {
    title: 'CONSUMER COMPLAINT PETITION',
    subtitle: 'Under Consumer Protection Act, 2019 | District Consumer Commission',
    icon: '🏛️',
    color: GOV_COLORS.saffron,
    slides: ['Cover', 'Complainant Details', 'Opposite Party Details', 'Facts of Complaint', 'Legal Grounds', 'Relief Sought', 'Documents & Verification']
  },
  'Bail Application': {
    title: 'APPLICATION FOR BAIL',
    subtitle: 'Under Section 437/438 CrPC | Honourable Court',
    icon: '⚖️',
    color: GOV_COLORS.navyBlue,
    slides: ['Cover', 'Applicant / Accused Details', 'FIR & Case Details', 'Grounds for Bail', 'Legal Arguments', 'Surety & Conditions', 'Prayer & Verification']
  }
};

// ── LOAD PPTXGENJS ───────────────────────────────────────────────
function loadPptxGenJS() {
  return new Promise((resolve, reject) => {
    if (window.PptxGenJS) { resolve(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
    s.onload = () => { console.log('✅ PptxGenJS loaded'); resolve(); };
    s.onerror = () => reject(new Error('Failed to load PptxGenJS'));
    document.head.appendChild(s);
  });
}

// ── GENERATE PPT ─────────────────────────────────────────────────
async function generatePPT() {
  if (!lastGeneratedText || !lastGeneratedTitle) {
    showToast('Please generate the document text first, then download as PPT.', 'error');
    return;
  }

  const btn = document.getElementById('ppt-dl-btn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="btn-spinner"></span> Creating PPT...'; }

  try {
    await loadPptxGenJS();

    // Get AI-structured content for slides
    showToast('🤖 Kanon AI is structuring your slides...');
    const template = DOC_PPT_TEMPLATES[lastGeneratedTitle] || DOC_PPT_TEMPLATES['Legal Notice'];

    const aiPrompt = `You are a government document presentation specialist. Structure this Indian legal document into presentation slides.

Document Type: ${lastGeneratedTitle}
Document Content:
${lastGeneratedText.slice(0, 3000)}

Create exactly ${template.slides.length} slides with these titles: ${template.slides.join(', ')}

Respond ONLY with valid JSON array (no markdown):
[
  {
    "slideTitle": "Exact slide title from the list",
    "heading": "Bold heading for this slide",
    "bullets": ["Point 1", "Point 2", "Point 3"],
    "keyInfo": "One important highlighted text",
    "footer": "Relevant legal section or date"
  }
]`;

    const rawAI = await callKanonAI(aiPrompt);
    let slides;
    try {
      slides = JSON.parse(rawAI.replace(/```json|```/g, '').trim());
    } catch(e) {
      // Fallback: create basic slides from document text
      slides = createFallbackSlides(template, lastGeneratedText);
    }

    // Build PPT
    const prs = new PptxGenJS();
    prs.layout = 'LAYOUT_WIDE'; // 16:9, standard government format
    prs.title = template.title;
    prs.subject = 'Indian Legal Document — NyayaMitra';
    prs.company = 'Government of India — Ministry of Law and Justice';

    // ── SLIDE 1: COVER ────────────────────────────────────────────
    const cover = prs.addSlide();
    
    // Full background
    cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: template.color } });
    
    // Top stripe (Saffron)
    cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.4, fill: { color: GOV_COLORS.saffron } });
    
    // Bottom stripe (Green)
    cover.addShape(prs.ShapeType.rect, { x: 0, y: 6.9, w: '100%', h: 0.35, fill: { color: GOV_COLORS.green } });
    
    // White center panel
    cover.addShape(prs.ShapeType.rect, { x: 0.8, y: 1.2, w: 8.4, h: 4.8, fill: { color: 'FFFFFF' }, line: { color: GOV_COLORS.gold, width: 2 } });

    // Government of India text
    cover.addText('GOVERNMENT OF INDIA', {
      x: 0, y: 0.45, w: '100%', fontSize: 11, bold: true,
      color: 'FFFFFF', align: 'center', fontFace: 'Arial'
    });
    cover.addText('MINISTRY OF LAW AND JUSTICE', {
      x: 0, y: 0.7, w: '100%', fontSize: 9,
      color: 'FFFFFF', align: 'center', fontFace: 'Arial'
    });

    // Emblem circle placeholder
    cover.addShape(prs.ShapeType.ellipse, { x: 4.1, y: 1.35, w: 1.0, h: 1.0, fill: { color: template.color }, line: { color: GOV_COLORS.gold, width: 2 } });
    cover.addText('⚖', { x: 4.1, y: 1.55, w: 1.0, fontSize: 28, color: 'FFFFFF', align: 'center', fontFace: 'Arial' });

    // Document Title
    cover.addText(template.title, {
      x: 0.9, y: 2.45, w: 8.2, h: 0.7,
      fontSize: 22, bold: true, color: template.color,
      align: 'center', fontFace: 'Arial'
    });

    // Gold divider
    cover.addShape(prs.ShapeType.line, {
      x: 1.5, y: 3.2, w: 7.0, h: 0,
      line: { color: GOV_COLORS.gold, width: 2 }
    });

    // Subtitle
    cover.addText(template.subtitle, {
      x: 0.9, y: 3.3, w: 8.2, h: 0.4,
      fontSize: 11, color: GOV_COLORS.midGray,
      align: 'center', fontFace: 'Arial', italic: true
    });

    // Date and Ref
    const today = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });
    cover.addText(`Date: ${today}    |    Ref. No.: _______________/${new Date().getFullYear()}`, {
      x: 0.9, y: 3.8, w: 8.2, h: 0.3,
      fontSize: 10, color: GOV_COLORS.midGray, align: 'center', fontFace: 'Arial'
    });

    // NyayaMitra footer
    cover.addText('Generated by NyayaMitra AI | nyayamitra.in | For Official Use', {
      x: 0, y: 7.0, w: '100%', fontSize: 8,
      color: 'FFFFFF', align: 'center', fontFace: 'Arial'
    });

    // ── CONTENT SLIDES ────────────────────────────────────────────
    slides.forEach((slide, idx) => {
      const s = prs.addSlide();

      // Top header bar
      s.addShape(prs.ShapeType.rect, {
        x: 0, y: 0, w: '100%', h: 0.65,
        fill: { color: template.color }
      });

      // Saffron left accent
      s.addShape(prs.ShapeType.rect, {
        x: 0, y: 0, w: 0.08, h: '100%',
        fill: { color: GOV_COLORS.saffron }
      });

      // Slide title in header
      s.addText(slide.slideTitle || slide.heading || `Slide ${idx + 2}`, {
        x: 0.2, y: 0.1, w: 8.5, h: 0.45,
        fontSize: 14, bold: true, color: 'FFFFFF', fontFace: 'Arial'
      });

      // Slide number
      s.addText(`${idx + 2}/${slides.length + 1}`, {
        x: 8.8, y: 0.15, w: 1.0, h: 0.35,
        fontSize: 9, color: 'FFFFFF', align: 'right', fontFace: 'Arial'
      });

      // Ministry watermark (light)
      s.addText('GOVERNMENT OF INDIA', {
        x: 0, y: 3.0, w: '100%',
        fontSize: 48, color: 'F0F0F0', align: 'center',
        fontFace: 'Arial', bold: true,
        transparency: 90
      });

      // Heading
      if (slide.heading) {
        s.addShape(prs.ShapeType.rect, {
          x: 0.15, y: 0.75, w: 9.6, h: 0.42,
          fill: { color: 'F0F4FF' }, line: { color: template.color, width: 1 }
        });
        s.addText(slide.heading, {
          x: 0.25, y: 0.8, w: 9.4, h: 0.35,
          fontSize: 12, bold: true, color: template.color, fontFace: 'Arial'
        });
      }

      // Key Info box (highlighted)
      if (slide.keyInfo) {
        s.addShape(prs.ShapeType.rect, {
          x: 0.15, y: 1.25, w: 9.6, h: 0.5,
          fill: { color: GOV_COLORS.saffron + '20' },
          line: { color: GOV_COLORS.saffron, width: 1.5 }
        });
        s.addText('📌 ' + slide.keyInfo, {
          x: 0.25, y: 1.3, w: 9.4, h: 0.4,
          fontSize: 11, bold: true, color: GOV_COLORS.darkGray, fontFace: 'Arial'
        });
      }

      // Bullet points
      const bulletStart = slide.keyInfo ? 1.85 : 1.25;
      if (slide.bullets && slide.bullets.length) {
        const bulletItems = slide.bullets.map(b => ({
          text: b, options: {
            fontSize: 11, color: GOV_COLORS.darkGray, fontFace: 'Arial',
            bullet: { type: 'number', color: template.color }
          }
        }));
        s.addText(bulletItems, {
          x: 0.3, y: bulletStart, w: 9.3,
          h: Math.min(slide.bullets.length * 0.55 + 0.3, 4.5),
          fontSize: 11, color: GOV_COLORS.darkGray, fontFace: 'Arial',
          lineSpacingMultiple: 1.4
        });
      }

      // Footer
      s.addShape(prs.ShapeType.rect, {
        x: 0, y: 6.9, w: '100%', h: 0.35,
        fill: { color: template.color }
      });
      s.addText(
        `${template.title}  |  ${slide.footer || today}  |  CONFIDENTIAL`,
        {
          x: 0.15, y: 6.95, w: 9.7, h: 0.25,
          fontSize: 7, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
        }
      );
    });

    // ── FINAL SLIDE: VERIFICATION ─────────────────────────────────
    const fin = prs.addSlide();
    fin.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: 'F9F9F9' } });
    fin.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 0.65, fill: { color: template.color } });
    fin.addText('VERIFICATION & SIGNATURES', {
      x: 0.2, y: 0.1, w: 8.5, h: 0.45,
      fontSize: 14, bold: true, color: 'FFFFFF', fontFace: 'Arial'
    });

    // Signature boxes
    const sigBoxes = [
      { label: 'Complainant / Applicant', x: 0.4 },
      { label: 'Witness 1', x: 3.6 },
      { label: 'Notary / Attesting Officer', x: 6.8 }
    ];
    sigBoxes.forEach((sb, i) => {
      fin.addShape(prs.ShapeType.rect, {
        x: sb.x, y: 1.5, w: 2.8, h: 2.2,
        fill: { color: 'FFFFFF' }, line: { color: GOV_COLORS.darkGray, width: 1 }
      });
      fin.addText('(Signature)', {
        x: sb.x, y: 2.1, w: 2.8, h: 0.3,
        fontSize: 9, color: 'BBBBBB', align: 'center', fontFace: 'Arial', italic: true
      });
      fin.addShape(prs.ShapeType.line, {
        x: sb.x + 0.2, y: 3.5, w: 2.4, h: 0,
        line: { color: GOV_COLORS.darkGray, width: 1 }
      });
      fin.addText(sb.label, {
        x: sb.x, y: 3.6, w: 2.8, h: 0.25,
        fontSize: 9, color: GOV_COLORS.darkGray, align: 'center', fontFace: 'Arial', bold: true
      });
    });

    fin.addText('Date: ________________   Place: ________________   Seal:', {
      x: 0.4, y: 4.4, w: 9.2, h: 0.3,
      fontSize: 10, color: GOV_COLORS.darkGray, fontFace: 'Arial'
    });

    fin.addShape(prs.ShapeType.rect, { x: 6.0, y: 4.3, w: 2.5, h: 1.8, fill: { color: 'FFFFFF' }, line: { color: GOV_COLORS.darkGray, width: 1 } });
    fin.addText('OFFICIAL SEAL', { x: 6.0, y: 5.1, w: 2.5, h: 0.3, fontSize: 8, color: 'BBBBBB', align: 'center', fontFace: 'Arial', italic: true });

    fin.addShape(prs.ShapeType.rect, {
      x: 0, y: 6.9, w: '100%', h: 0.35,
      fill: { color: template.color }
    });
    fin.addText('Generated by NyayaMitra AI — Official Legal Document System | Government of India Format', {
      x: 0, y: 6.95, w: '100%', h: 0.25,
      fontSize: 7, color: 'FFFFFF', align: 'center', fontFace: 'Arial'
    });

    // Save
    const filename = (lastGeneratedTitle || 'LegalDoc').replace(/\s+/g, '_') + '_Official.pptx';
    await prs.writeFile({ fileName: filename });
    showToast('✅ Official PPT downloaded successfully!');

  } catch(e) {
    showToast('PPT Error: ' + e.message, 'error');
    console.error(e);
  }

  if (btn) {
    btn.disabled = false;
    btn.innerHTML = '<span>📊</span> Download as PPT';
  }
}

// ── FALLBACK SLIDE CREATOR ────────────────────────────────────────
function createFallbackSlides(template, docText) {
  const lines = docText.split('\n').filter(l => l.trim()).slice(0, 60);
  const chunkSize = Math.ceil(lines.length / (template.slides.length - 1));
  return template.slides.slice(1).map((title, i) => {
    const chunk = lines.slice(i * chunkSize, (i + 1) * chunkSize);
    return {
      slideTitle: title,
      heading: title,
      bullets: chunk.slice(0, 6).map(l => l.trim().slice(0, 120)),
      keyInfo: chunk[0] ? chunk[0].slice(0, 100) : '',
      footer: new Date().toLocaleDateString('en-IN')
    };
  });
}

// ── AI CHAT WITH KANON ────────────────────────────────────────────
async function sendChatWithKanon(userMsg, chatHistory, systemPrompt) {
  // Build messages array for Kanon API
  const messages = [
    { role: 'system', content: systemPrompt },
    ...chatHistory.map(h => ({
      role: h.role === 'model' ? 'assistant' : 'user',
      content: h.parts[0].text
    })),
    { role: 'user', content: userMsg }
  ];

  try {
    const response = await fetch('https://api.kanonaigc.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${KANON_API_KEY}`
      },
      body: JSON.stringify({
        model: 'gemini-2.0-flash',
        messages: messages,
        temperature: 0.3,
        max_tokens: 2000,
        response_format: { type: 'json_object' }
      })
    });

    if (response.ok) {
      const data = await response.json();
      const text = data?.choices?.[0]?.message?.content;
      if (text) return text;
    }
  } catch(e) {
    console.warn('Kanon chat failed, falling back:', e.message);
  }

  // Fallback to direct Gemini
  return await callGeminiDirect(userMsg);
}

console.log('✅ NyayaMitra PPT Generator + Kanon AI loaded');
