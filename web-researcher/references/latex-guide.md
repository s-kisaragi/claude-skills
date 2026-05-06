# LaTeX Document Guide

## Engine Selection

| Engine | Use Case | Japanese Support |
|:-------|:---------|:-----------------|
| lualatex | Recommended for Japanese | Native via luatexja |
| xelatex | Alternative for CJK | Via xeCJK package |
| pdflatex | English-only documents | Not recommended for Japanese |

## Template Selection

- **English report**: Use `assets/report-template.tex`
- **Japanese report**: Use `assets/report-template-ja.tex` (requires lualatex)

## Common Patterns

### Figures
```latex
\begin{figure}[H]
  \centering
  \includegraphics[width=0.8\textwidth]{image.png}
  \caption{Description}
  \label{fig:label}
\end{figure}
```

### Tables
```latex
\begin{table}[H]
  \centering
  \caption{Description}
  \begin{tabular}{lrr}
    \toprule
    Item & Value & Note \\
    \midrule
    A & 100 & text \\
    B & 200 & text \\
    \bottomrule
  \end{tabular}
  \label{tab:label}
\end{table}
```

### Code blocks
```latex
\begin{lstlisting}[language=Python,caption={Description}]
def example():
    return "hello"
\end{lstlisting}
```

## Building PDF

Run the build script:
```bash
bash ~/.claude/skills/web-researcher/scripts/build_pdf.sh input.tex [output_dir]
```

The script automatically:
1. Detects available LaTeX engine
2. Runs bibliography processing if needed
3. Runs multiple passes for cross-references

## Troubleshooting

- **Missing packages**: `tlmgr install <package-name>`
- **Japanese not rendering**: Ensure lualatex is used, not pdflatex
- **Bibliography not showing**: Ensure biber/bibtex is installed and .bib file is referenced
