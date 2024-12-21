<img src="perpop.png" alt="Andel Projects Limited" width="400">

# Conclusions on Critical Path Drag (Method A vs. Method B)

## 1. Introduction

**Critical Path Method** (CPM) scheduling has long been a cornerstone of project management. One of its key concepts **Critical Path Drag** was popularized by **Stephen Devaux** to measure how much each critical activity is “dragging” or extending the overall project finish date.

Knowing drag values helps project managers:

- Pinpoint which tasks offer the _best leverage_ if shortened (crashed)
- Identify where resources or funds for schedule compression should be allocated

Over time, researchers and practitioners have recognized that **two broad approaches** exist to compute Drag:

1. **Method A**: The “Zero‐Duration Thought Experiment.”

   - Temporarily set a critical task’s duration to zero.
   - Recalculate the project finish date.
   - The difference in finish dates (before vs. after) = that task’s “drag.”

   This was the approach taken by **Alex Lyaschenko** in developing his macro which he kindly shared on LinkedIn.

2. **Method B**: The “Concurrency‐ or Logic‐Based Approach.”

   - Conceptually:  
    <img src="equation" alt="CP Drag Formula" width="400">
   - In code, often done by analyzing overlaps, total floats, and re‐running partial or incremental schedules.

   Andy Forrester addopted this approach.

## Questions following comparison testing

- Why might **Method A** yield a **different drag value** than Method B, even if the task appears to be on the critical path without obvious parallel tasks?
- How do **partial reductions** (not setting a task’s duration to zero, but e.g. reducing it by 1 or 2 days) sometimes cause other paths to become critical, changing the overall drag?
- **How** can we implement code in Microsoft Project (via VBA) to accurately capture drag, including cost/benefit analysis?
- What about “**reverse drag**” scenarios (a.k.a. “negative drag”), or complex dependencies like SS/FF/SF that can push or pull early‐project dates?

---

## Problem

### Observed Discrepancy in Drag Values

A key problem arose when it noticed that:

- **Method A** (zero‐duration) yielded a **2‐day** reduction in overall project finish if a certain 4‐day task were removed entirely.
- **Method B** (the concurrency check) seemed to report a **4‐day** drag for the same task.

On the surface, it looked contradictory. After all, if the task is genuinely “on the critical path,” one might expect removing it entirely to save the full 4 days in the overall schedule. Yet it only saved 2 days.

### Deeper Investigation: “Parallel Tasks” or <font color = "red">“Critical Path Shifts”</font>

Upon further investigation, it became clear that even a seemingly **linear** chain can hide subtleties:

1. **When you remove a task entirely**:

   - The network logic can _re‐sequence_ or reveal a near‐critical path that only becomes controlling after a chunk of the main critical path disappears.
   - This can limit the net improvement to only 2 days, not 4.

2. **When analyzing concurrency** in a single pass (Method B as originally coded, comparing total slack of overlapping tasks):
   - **The code sees _no concurrent tasks_ in that moment, so it reports the _full_ 4 days as “drag.”**
   - It was **_not recalculating_** the schedule after partial reductions to see if something else would become critical earlier.

Thus, the discrepancy arose from **<font color="red">_critical‐path switching_** </font>once certain durations fell below a threshold.

---

## Underlying Causes of the Discrepancy

1. **Network Logic & Partial Reductions**

   - Schedules are often nonlinear and can contain multiple convergent or divergent paths, resource constraints, or near‐critical paths.
   - A partial shortening of an activity might or might not shift the critical path.

2. **Method A vs. Method B**

   - **Method A** is a _pure “What if the task was zero?”_ scenario—simple, but does not capture how _partial_ reductions might change concurrency.
   - **Method B** attempts to measure concurrency _in the existing network_, but typically needs **iterative** recalculations (for partial day reductions) to capture the dynamic nature of critical paths.

3. **Resource Leveling, Lags, and Constraints**
   - Even if the initial plan looks “simple,” real projects frequently have constraints or resource leveling that reorder tasks in ways that do not appear in a single pass.
   - Different dependency types (SS, FF, SF) complicate calculations further.

---

## Proposal to Retain & Update Method B

Despite Method A’s simplicity, it is proposed that a community make a decision on the way forward to <font color="red">Standardise</font> on **retaining and refining Method B** the concurrency/logic approach—because:

1. **Incremental Accuracy**
   Ability to see the effect of _partially_ reducing a critical task’s duration in increments.

   - Method B, when done iteratively, can recalculate the entire schedule each time we shave off 1 day (or half a day) from a task. It detects any new concurrency or critical path changes at each step.

2. **Future‐Proofing**

   - It is anticipated that once a concensus is reached then adding **reverse drag** or “negative drag” features will be undertaken and published for review soon after .
   - **Method B’s step‐by‐step recalculations more readily extend to analyzing how a task might pull forward or delay _upstream_ tasks in scenarios with Start‐to‐Start, Finish‐to‐Finish, or Start‐to‐Finish links.**

3. **Enhanced Cost/Benefit Analysis**
   - Method B can do dynamic cost calculations each time a partial reduction is tested.
   - This is valuable for managers deciding which tasks to crash for maximum ROI in schedule compression.

**<font color                                                                                                           = "red">Important</font>**
Any solution adopt now is primarily geared toward _straightforward Finish‐to‐Start (FS) relationships_. <font color = "red">Tasks that use SS, FF, or SF relationships—or tasks subject to resource leveling—may not behave in a purely linear manner. While Method B _can_ handle these complexities with repeated recalculations, the code examples we’ve provided are simplified for an FS‐only or predominantly FS environment. </font>

Future enhancements will include:

- Accounting for partial overlaps or lags in SS, FF, SF links,
- Additional logic to detect “reverse drag” (i.e., when changes in successor tasks can push or pull predecessor tasks or early milestones).

---

## Summary of the Findings

1. **Discrepancies between Method A & Method B** are natural and expected when partial duration changes cause the critical path to shift or when near‐critical paths become controlling.
2. **Method A** is quick and easy but focuses only on a “zero‐duration” scenario, which can mask partial concurrency.
3. **Method B** is more powerful but requires iterative recalculations to handle partial reductions, concurrency detection, and advanced dependency types.
4. **For “negative” or “reverse” drag** (i.e., when you want to see how a task affects _upstream_ or earlier milestones), Method B’s iterative logic approach is again more suitable, because Method A typically only focuses on the final completion milestone.
5. This paper **communicates the desire for concensus within the community** such that retention and an a similar approach can be adopted for further updating **Method B** to facilitate negative drag capability. This asssumes the stepwise concurrency approach can handle them in a logically consistent manner.

---

## Conclusion

This informal paper outlines:

- **The problem** of conflicting drag results when comparing a zero‐duration thought experiment to a concurrency‐based single‐pass approach,
- **Why partial reductions** in task duration can reveal different concurrency or critical‐path structures,
- **How** iterative (Method B) calculations resolve these discrepancies by recalculating the entire network logic after each partial change,
- **The rationale** for continuing with Method B for future expansions, including negative drag and more advanced dependency types, and
- **A note** that the provided approach currently assumes relatively simple, _Finish‐to‐Start_ relationships, though it can be extended to more complex dependencies with additional coding logic and repeated recalculations.

Ultimately, **Method B** is more robust for real‐world scenarios where concurrency, resource constraints, and partial crashing are common. It demands more computational work but provides a more accurate reflection of how actual project networks behave under incremental schedule adjustments.

My thanks go to all contributers on the LinkedIn thread and especially to **Alex Lyaschenko** for his macro and feedback
