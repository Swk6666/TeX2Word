---
title: "TeX2Word Markdown 示例"
---

# 第 1 章 简介

这是一份包含中英文正文、行内公式、独立公式与表格的 Markdown 示例文档。支持标准 $a^2 + b^2 = c^2$ 语法、LLM 常见的 \(E = mc^2\) 行内公式，以及带空格的行内公式 $ \alpha + \beta = \gamma $，同时不会把普通引用 \[1\] 误判为公式。

\[
\int_0^1 x^2 \,\mathrm{d}x = \frac{1}{3} \tag{1-1}
\]

\begin{equation}
F = -k x \label{eq:hooke}
\end{equation}

由式 \eqref{eq:hooke} 可知胡克定律与线性刚度相关。代码块中的公式不会被误处理：

```python
expr = r"\(a + b\) and $x^2$"
```

## 1.1 参数表

| 参数 | 表达式 | 单位 |
| --- | --- | --- |
| 刚度 | $k_1$ | $\mathrm{N/m}$ |
| 频率 | $f_0 = \frac{1}{2\pi}\sqrt{\frac{k}{m}}$ | $\mathrm{Hz}$ |

<table>
<thead><tr><th>指标</th><th>公式</th></tr></thead>
<tbody><tr><td>能量</td><td>$U = \frac{1}{2}kx^2$</td></tr></tbody>
</table>
