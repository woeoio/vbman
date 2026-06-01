# VBMAN License Usage Guide

This document helps you quickly understand the licensing requirements of the VBMAN project and when you need to pay for a commercial license.

---

## Core Points of the Existing License

### Base License: GPL-3.0

VBMAN is licensed under the **GNU General Public License v3.0 (GPL-3.0)**.

## One-Sentence Summary

> **Using compiled DLL → Completely Free**  
> **Modifying source code and open-sourcing → Completely Free**  
> **Using DLL is free; modifying source code and using it closed-source requires payment.**

---

## Detailed Explanation

### 1. Binary Files (Permanently Free)

VBMAN compiled binary files (`.dll`, `.exe`, etc.) are **permanently free** with no usage restrictions:

- ✅ Personal projects — Free
- ✅ Commercial software directly referencing VBMAN DLL — Free
- ✅ Internal company tool development — Free
- ✅ Open source projects — Free

### 2. Source Code Usage

#### Case 1: Free Usage (Complying with GPL)

If you meet any of the following conditions, you can use the source code for free:

1. **Personal learning/research** — Modifying for your own use, not distributed externally
2. **Modified and open-sourced** — If you modify VBMAN and distribute it (including selling), you must open-source the entire project under a GPL-compatible license

**Core GPL Requirements:**

- Must include LICENSE file when distributing
- Modified works must also be open-sourced under GPL
- Retain copyright notices and disclaimers

#### Case 2: Requires Commercial License

Only one situation requires payment:

> **You want to use VBMAN source code in a closed-source commercial project and do not want to open-source your own code.**

Common scenarios:

| Scenario | Description |
| ------------------------------------------ | ------------ |
| Modified VBMAN source code but don't want to open-source modifications | Requires commercial license |
| Integrated VBMAN source code into closed-source product for sale | Requires commercial license |
| SaaS service using modified VBMAN source code but closed-source deployment | Requires commercial license |

---

## Decision Flowchart

```
┌─────────────────────────────────────┐
│     Do you want to use VBMAN?       │
└──────────────────┬──────────────────┘
                   │
         ┌─────────┴─────────┐
         ▼                   ▼
   ┌──────────┐        ┌──────────┐
   │  Use DLL │        │ Use Source │
   └────┬─────┘        └────┬─────┘
        │                   │
        ▼                   ▼
   ┌──────────┐        ┌──────────────┐
   │ Free     │        │ Distribute?  │
   │ No License│       └──────┬───────┘
   └──────────┘               │
                      ┌───────┴───────┐
                      ▼               ▼
                ┌──────────┐    ┌──────────┐
                │Closed    │    │Open      │
                │Source    │    │Source    │
                └────┬─────┘    └────┬─────┘
                     │               │
                     ▼               ▼
               ┌──────────┐    ┌──────────┐
               │Purchase  │    │Completely│
               │Commercial│    │Free      │
               │License   │    │GPL       │
               └──────────┘    └──────────┘
```

## Summary

### Definition of Payment Situations

**1. Binary Files (DLL, EXE, etc.)**

- Permanently free
- No usage restrictions
- Whether for personal or commercial use

**2. Source Code Usage - Individual Users**

- Free to use
- But must comply with GPL license requirements:
  - Must include LICENSE file when distributing
  - Modified works must also be open-sourced under GPL-compatible license

**3. Source Code Usage - Commercial Scenarios**
There are two situations here:

**Scenario A: Open Source Commercial Use (Free)**

- If your commercial project is based on modifications or secondary development of this project
- And **when distributing/selling**, you **open-source the entire project's source code** under a GPL-compatible license
- In this case, **no payment is required**

**Scenario B: Closed Source Commercial Use (Requires Payment)**

- If you want to use this project's source code in a **closed-source commercial project**
- That is, you don't want to open-source your own code
- You must contact the author to purchase a commercial license

### Specific Scenarios Requiring Payment

1. Commercial software company developing closed-source products using VBMAN source code
2. Internal closed-source enterprise systems using VBMAN source code and not wanting to follow GPL open-source requirements
3. SaaS service providers building services on closed-source code

### Scenarios Not Requiring Payment

1. Directly using compiled DLL/EXE files, whether for personal or commercial use
2. Personal learning and research using source code
3. Projects based on source code development that are willing to open-source modified code
4. Any project following GPL open-source license

The core of commercial licensing lies in how the source code is used, not the deployment of binary files. The developer provides flexible usage paths, ensuring both open-source spirit and legal avenues for commercial applications. The key is whether the user is willing to bear the open-source obligation or choose to purchase a license for closed-source use.

---

## Commercial License Information

For closed-source commercial licensing, please contact:

- **Author**: Deng Wei (邓伟)
- **Website**: https://a-vi.com
- **Repository**: https://gitcode.com/woeoio/vbman

Commercial license fees are negotiated based on specific usage scenarios and scope.

---

## Frequently Asked Questions (FAQ)

**Q1: I use VBMAN to develop internal tools at my company. Do I need to pay?**  
A: No. Using the DLL directly is completely free.

**Q2: I modified VBMAN's source code, made it into a product and sold it to others, but I don't want to open-source it. What should I do?**  
A: You need to purchase a commercial license. Otherwise, it violates the GPL license.

**Q3: I developed an open-source software using VBMAN. Can I charge for it?**  
A: Yes. The GPL license allows you to sell for a fee, but you must provide the source code to customers.

**Q4: Is VBMAN2 under the same licensing policy?**  
A: VBMAN2 is not open-sourced, only binary files are provided, and they are permanently free to use.

---

## Full Legal Text

For detailed terms, please refer to the [LICENSE](LICENSE) file.
