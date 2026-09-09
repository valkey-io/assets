# The Vertical Scaling Wall — deck contents

File: `slides/The Vertical Scaling wall.pptx`
31 slides, 16:9 (12192000 x 6858000 EMU). Same Valkey template as the community deck, but the layout indices differ (see bottom).

## Slide-by-slide

| # | Title | Layout | Content |
|---|-------|--------|---------|
| 1 | The Vertical Scaling wall, and how Valkey gets over it | Title + 2x Speaker | Madelyn Olson / Valkey Maintainer and Engineer. Second speaker placeholders empty. |
| 2 | Throughput is a cliff | Title + Cont | Line-with-markers chart. Latency (avg/p50/p90/p95/p99) vs offered load 10K→165K rps. Smooth climb to 140K, then everything pins at ~4.7-5.0ms at 150K+. |
| 3 | Steady state is a lie | Title + Cont | Line chart, 25 unlabeled points. `Cache throughput` (spiky diurnal, 8K→180K) against flat `Provisioned capacity (r6g.large)` at 150K. |
| 4 | Hybrid scaling model | Title + Cont | Bullets: vertical scaling (add CPUs; constrained by hardware; larger blast radius on failure/patching) and horizontal scaling (shard + route; more components; limited burst absorption). |
| 5 | Thread scalable command processing | Title + Cont | Diagram build 1/5. Single "Main Thread" timeline: System Epoll → Read Client Command → Parse Command → Execute Command → Send response. |
| 6 | Thread scalable command processing | Title + Cont | Build 2/5. Adds "IO Thread 1": read/parse/send move off the main thread, Execute stays on main, four handoff arrows between the two timelines. |
| 7 | Thread scalable command processing | Title + Cont | Build 3/5. Two request cycles pipelined across main + IO Thread 1. |
| 8 | Thread scalable command processing | Title + Cont | Build 4/5. Adds "IO Thread 2" with its own read/parse/send column. |
| 9 | Thread scalable command processing | Title + Cont | Build 5/5. Main thread is a solid column of Execute Command (5 blocks) fed by IO Thread 1 + IO Thread 2. The payoff frame. |
| 10 | Hybrid scaling model | Title + Cont | Repeat of slide 4 with the vertical-scaling bullets removed (horizontal only). |
| 11 | *(no title)* | Title + Cont | Two rounded shards side by side: "Shard 1 (Slots 0 – 8196)", "Shard 2 (Slots 8197 – 16343)", each Primary + Replica with replication arrow. |
| 12 | *(no title)* | Title + Cont | Three shards: Shard 1 (0 – 8196), Shard 3 (No slots), Shard 2 (8197 – 16383). New empty shard added, no slots yet. |
| 13 | *(no title)* | **Blank** | Same three shards after rebalance: Shard 1 (0 – 5765), Shard 3 (5766 – 10922), Shard 2 (10923 – 16383), plus two arrows showing slots migrating into Shard 3. |
| 14 | Valkey performance is bi=modal | Title + Cont | Scatter-lines: CPU Utilization (~50% → ~200%) and P50 Latency (~25ms → ~500ms) both step-change at the same point; flat `Throughput (constant)` reference at 2500. |
| 15 | A story about a customer | Title + Cont | Full-width screenshot (PNG). |
| 16 | Metastability | Title + Cont | Single scatter-line series `Healthy State`: ramps 0→75, plateaus, then collapses to 2 and stays there. |
| 17 | Stable system | Title Only | Custom diagram. Clients row of req boxes → cache layer banner "Cache UP" → 20-cell Backend Database row. Includes an overlaid layer of retry/retry² boxes (animation states stacked on the same coordinates). |
| 18 | Unstable system | Title Only | Same diagram, banner flipped to "✖ CACHE DOWN / COLD", clients row now req, req, req, req, retry, retry, retry², retry², retry², retry². |
| 19 | Throughput is a cliff | **DEFAULT** | Duplicate of slide 2's chart, built as a raw text box + chart instead of placeholders. Looks like a leftover. |
| 20 | Different system modes of Valkey | Title Only | Two-card comparison. State A: Steady state / Static topology / Full data. State B: Replication / Slot rebalancing / Empty cache. Card titles are literally "State A" and "State B". |
| 21 | How did this system fail | Title + Cont | Application → Primary (writes) and → Replica (reads), replication arrow, with a highlight oval drawn over the replica. |
| 22 | First in, never out | Title + Cont | Application → Kafka Queue (writes) and → Context Cache (writes); Consumer reads from both. Full picture. |
| 23 | First in, never out | Title + Cont | Same diagram with the Consumer's read arrows removed (build step). Appears *after* the complete version. |
| 24 | First in, never out | Title + Cont | Same as 22 (read arrows restored). |
| 25 | To proxy, or not to proxy | Title + Cont | Application → Proxy → Cache 1 / Cache 2 / Cache 3. |
| 26 | To proxy, or not to proxy | Title + Cont | Same plus two `req` boxes in flight on the Application → Proxy hop. |
| 27 | Valkey's architectural goals | Title + Cont | Two bullets: keep core execution simple while offloading expensive operations; make scaling operations seamless. |
| 28 | Takeaways | Divider | Section divider. |
| 29 | Think about failure modes | Title + Cont | What are your operating modes? What triggers stable→unstable? Can the system recover on its own? What breaks the feedback loop? |
| 30 | Some best practices | Title + Cont | Test partitions/cache failures/transitions; handle timeouts and trigger right behavior; scale proactively; measure the right bottlenecks. |
| 31 | *(Thank you)* | Thank You 2x | Madelyn Olson, @reconditerose, github.com/madolson. Second speaker placeholders empty. |

## Narrative arc as it currently stands

1. **Hook (2-3)** — throughput is a cliff, and steady state is a lie. You cannot provision for the peak.
2. **Scaling options (4)** — vertical vs horizontal, each with costs.
3. **Vertical answer (5-9)** — IO threading lets the main thread do nothing but execute.
4. **Horizontal answer (10-13)** — sharding, then adding a shard and rebalancing slots.
5. **Why it goes wrong (14-20)** — bimodal performance, a customer story, metastability, stable vs unstable system, the two system modes.
6. **Concrete failures (21-26)** — replica read failure, Kafka + context cache, proxy tradeoff.
7. **Close (27-31)** — architectural goals, takeaways, best practices.

## Problems I found

**Blockers**

- Slide 14 title typo: `bi=modal` → `bi-modal`.
- Slide 11 slot math is wrong: `Slots 8197 – 16343` (should be 16383, as on slides 12-13). The 8196/8197 split is also off by one from the conventional 0-8191 / 8192-16383.
- Slides 11, 12, 13 have no titles. Three diagram slides in a row with nothing on top.

**Structural**

- Slide 19 duplicates slide 2 and is the only slide on the `DEFAULT` layout with a hand-built text box instead of a title placeholder. Either delete it or merge it with slide 2.
- Slide 23 is a *reduced* version of slide 22, sitting after it, so the 22→23→24 sequence goes complete → partial → complete. 23 and 22 are almost certainly swapped.
- Slide 10 is slide 4 minus the vertical bullets. If it's meant as a build, 10 should come before 4; if it's a callback, it needs a different title so the audience knows why they're seeing it again.
- Slide 13 is on the `Blank` layout while 11 and 12 are on `Title + Cont`. Inconsistent, and it's why 13 has no footer/logo.

**Content gaps**

- Speaker notes on slides 5-9 are copy-pasted from a memory-efficiency slide in another deck. They talk about hash table pointer chasing and embedded allocations, which is unrelated to IO threading. No other slide has notes.
- Slide 20's cards are labeled "State A" and "State B". Name them what they are.
- Slides 2, 3, 14, 16 have no source citation or axis labels called out in the shapes. Charts carry real numbers with no attribution.
- Slide 15 ("A story about a customer") is a bare screenshot with no framing text.
- No agenda slide and no divider slides until slide 28, so the six-part arc above is invisible to the audience.
- Slide 1's second-speaker placeholders and slide 31's second-speaker placeholders are empty. Fine if solo, but they render as blank fields.

## Editing notes specific to this file

See `../slides-skill.md` for template mechanics. Notes that apply only to this deck:

- It has the fullest layout set of any deck here (17 layouts, including `DEFAULT` and `Blank`
  which no other deck has). Resolve layouts by name, not index.
- Slides 2, 3, 14, 16, 19 hold embedded charts. Slides 2 and 19 use category charts; 14 and 16 use
  XY scatter. Editing data means `chart.replace_data()` with the full series set, not text runs.
- Diagram slides (5-9, 11-13, 17-18, 21-26) are hand-built from `AUTO_SHAPE` shapes and
  connectors, and use `ROUNDED_RECTANGLE` throughout rather than the square-corner house style.
  Match the rounded style when editing these slides.
- Slides 17 and 18 stack overlapping `req`/`retry` boxes at identical coordinates as manual
  animation states, and reuse shape names (`Shape 3`, `Text 4` appear three times on slide 17).
  Select by position, not by name, on these two.
