# Excel Performance Analysis in a PowerPoint Template

## User story

A business team needs to turn a regional and channel performance workbook into an executive presentation while following the company's standard PowerPoint template. The analysis may change, but the background, palette, title hierarchy, and card positions must remain consistent.

This example recreates that workflow with a smaller fictional dataset. It preserves only the dimensions and reporting pattern needed to explain the process; it is not a scaled transformation of real business data.

## Demonstration input

Unit: CNY 10,000. Period: first half of a fictional reporting year.

| Region | Channel | Revenue budget | Actual revenue |
| --- | --- | ---: | ---: |
| Region A | Offline | 1,400 | 1,640 |
| Region A | Online | 800 | 760 |
| Region B | Offline | 1,600 | 1,800 |
| Region B | Online | 1,200 | 1,000 |
| Total | | 5,000 | 5,200 |

## Traceable findings

- Overall attainment: 5,200 ÷ 5,000 = 104.0%, or 200 above budget.
- Region B revenue: 1,800 + 1,000 = 2,800, exceeding Region A's 2,400.
- Offline revenue: 1,640 + 1,800 = 3,440, or 440 above the 3,000 budget, for approximately 114.7% attainment.
- Online revenue: 760 + 1,000 = 1,760, or 240 below the 2,000 budget, for 88.0% attainment.

These figures support variance reporting only. They do not prove that conversion, inventory, media spend, or any other factor caused the gap; causal analysis requires additional evidence.

## Template mapping

The overview page maps its title and summary to the overall findings. The four existing red-and-gold cards hold revenue, attainment, regional contribution, and the channel gap. The channel page reuses the original template's four outlined icon cards for offline revenue, online revenue, offline uplift, and the online shortfall.

The promotional hero was assembled with an image-generation tool and may contain scaling or redrawn details. `images/template-overview.en.png`, `images/report-overview.en.png`, and `images/report-channels.en.png` are English-localized explanatory visuals derived from direct slide renders. The original Chinese direct renders remain available without the `.en` suffix. Neither set is native PowerPoint acceptance evidence or a basis for claiming "90% similarity."

The original Excel workbook, complete PowerPoint template, final report containing real data, embedded fonts, and internal working records are not published.
