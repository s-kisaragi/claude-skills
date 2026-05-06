# Citation & Source Management Guide

## Markdown Citations

### Inline style
Use footnote-style references at the bottom of the document:

```markdown
According to recent research[^1], the approach has shown promising results[^2].

[^1]: Author Name, "Article Title", Publication, Year. URL
[^2]: Author Name, "Article Title", Publication, Year. URL
```

### Reference section
Always include a `## References` section at the end:

```markdown
## References

1. Author Name. "Article Title." *Publication Name*, Year. [Link](https://example.com)
2. Author Name. "Article Title." *Publication Name*, Year. [Link](https://example.com)
```

## LaTeX Citations

### Using biblatex
Create a `references.bib` file alongside the `.tex` file:

```bibtex
@online{key2024,
  author  = {Author Name},
  title   = {Article Title},
  year    = {2024},
  url     = {https://example.com},
  urldate = {2024-01-15},
}

@article{paper2024,
  author  = {First Author and Second Author},
  title   = {Paper Title},
  journal = {Journal Name},
  year    = {2024},
  volume  = {1},
  pages   = {1--10},
  doi     = {10.xxxx/xxxxx},
}
```

Then in `.tex`:
```latex
\addbibresource{references.bib}
% ... in text:
As shown by \cite{key2024}, ...
% ... at end:
\printbibliography
```

## Source Quality Assessment

When researching, evaluate sources by:

1. **Recency** - prefer sources from the last 2 years for rapidly evolving topics
2. **Authority** - official docs > peer-reviewed > reputable blogs > forums
3. **Consistency** - cross-reference claims across multiple sources
4. **Primary vs secondary** - prefer primary sources when available

## Source Priority (highest to lowest)

| Priority | Source Type | Example |
|:---------|:-----------|:--------|
| 1 | Official documentation | docs.python.org, RFC documents |
| 2 | Peer-reviewed papers | arXiv, IEEE, ACM |
| 3 | Authoritative tech blogs | Engineering blogs from major companies |
| 4 | Community knowledge | Stack Overflow (high-voted), GitHub discussions |
| 5 | General web | Blog posts, tutorials |
