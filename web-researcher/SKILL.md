---
name: web-researcher
description: Research topics from the web, academic papers, and technical documentation, then produce well-structured documents with full source citations. Use when asked to investigate, research, survey, or write reports on any topic. Supports output as Markdown, LaTeX, or PDF. Triggers on requests like "調べて", "リサーチして", "レポートにまとめて", "research", "investigate", "write a report on".
---

# Web Researcher

Conduct thorough web research and produce publication-ready documents with proper citations.

## Workflow

### Phase 1: Scope Definition

Clarify with the user before starting research:

1. **Topic & scope** — What to research, how deep
2. **Output format** — Markdown (default), LaTeX, or LaTeX→PDF
3. **Language** — Japanese or English (affects template choice)
4. **Audience** — Technical depth and tone

If the user's request is clear enough, skip clarification and proceed.

### Phase 2: Research

Conduct research using multiple source types in priority order:

1. **Official documentation** — authoritative primary sources
2. **Academic papers** — arXiv, IEEE, ACM, Google Scholar
3. **Technical blogs & articles** — engineering blogs, reputable tech publications
4. **Community knowledge** — Stack Overflow, GitHub discussions

For each source:
- Record the URL, title, author, date
- Extract key claims and data points
- Cross-reference important claims across multiple sources
- Note any contradictions or uncertainties

Use `WebSearch` for broad discovery, `WebFetch` for extracting detailed content from specific pages.

**Research depth guideline:**
- Quick overview: 3–5 sources
- Standard report: 8–15 sources
- Deep dive: 15+ sources, cross-referenced

### Phase 3: Document Generation

#### Markdown Output (default)

Structure the document as:

```markdown
# Title

## Summary
Brief overview (2-3 sentences)

## Table of Contents
(for longer documents)

## Section 1: ...
Content with inline citations[^1]

## Section 2: ...
...

## Conclusion
Key takeaways

## References
[^1]: Author. "Title." Publication, Year. [Link](URL)
[^2]: ...
```

Rules:
- Every factual claim must have a citation
- Use footnote-style references `[^N]`
- Include a complete `## References` section at the end
- Write the output file to the user's working directory

#### LaTeX Output

1. Copy the appropriate template from assets:
   - English: `assets/report-template.tex`
   - Japanese: `assets/report-template-ja.tex` (requires lualatex)
2. Replace placeholder values (REPORT\_TITLE, REPORT\_AUTHOR, etc.)
3. Populate sections with researched content
4. Generate `references.bib` for citations
5. Uncomment `\addbibresource` and `\printbibliography` lines

For LaTeX formatting details, see [references/latex-guide.md](references/latex-guide.md).

#### PDF Output

After generating the `.tex` file, build PDF:

```bash
bash ~/.claude/skills/web-researcher/scripts/build_pdf.sh <file.tex>
```

If no LaTeX engine is installed, inform the user:
```
LaTeX is not installed. To enable PDF generation:
  brew install --cask mactex    # full (~4GB)
  brew install --cask basictex  # minimal (~100MB)
```

### Phase 4: Review

After generating the document:
1. Verify all citations have corresponding entries in References
2. Check for logical flow between sections
3. Confirm no unsupported claims remain
4. Present a brief summary of the document to the user

## Citation Requirements

**All sources must be cited.** See [references/citation-guide.md](references/citation-guide.md) for formatting details.

- Every factual claim needs a source
- URLs must be included for all web sources
- For academic papers, include DOI when available
- Access date should be noted for web sources

## Resources

### scripts/
- `build_pdf.sh` — Compile LaTeX to PDF (auto-detects engine, handles bibliography)

### references/
- [citation-guide.md](references/citation-guide.md) — Citation formatting for Markdown and LaTeX
- [latex-guide.md](references/latex-guide.md) — LaTeX template usage and common patterns

### assets/
- `report-template.tex` — English report template
- `report-template-ja.tex` — Japanese report template (lualatex)
