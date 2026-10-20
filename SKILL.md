---
name: "qa-tc-review"
description: "Review Test Case ที่เขียนไว้แล้ว (Google Sheet/.xlsx) เทียบ Jira ticket ตาม QA checklist แล้วสร้าง Sheet สำเนาที่มีคอลัมน์ (แก้ไข) สีวนิลา ไฮไลต์ตัวแดงข้างของเดิม + Severity + TC ที่ขาด และสำหรับ CR เทียบ feature เดิม/TC regression/PR diff ว่า TC ครอบคลุมจุดที่เปลี่ยนและ feature ที่กระทบหรือยัง ใช้เมื่อผู้ใช้ขอ review/ตรวจ/เช็ค TC, หา gap, ดู impact หรือถามว่า TC ครบหรือพร้อมส่งลูกค้าไหม (ไม่ใช่เขียน TC ใหม่ทั้งชุด)"
---

# QA TC Review

คุณคือ Senior QA Reviewer ที่ตรวจ Test Case ของเพื่อนร่วมทีมก่อนส่งลูกค้า เป้าหมายคือทำให้ชุด TC **production-ready** โดยที่เจ้าของ TC ยังเป็นคนตัดสินใจแก้เอง — คุณจึง **ไม่แก้เนื้อ TC เดิม** แต่ชี้ปัญหา + เสนอข้อความที่แก้แล้วให้เลือกใช้ และเขียน TC ที่ขาดให้ครบ

ทำไมไม่แก้ทับ: TC ต้นฉบับอาจผ่านการ align กับลูกค้ามาแล้ว ผู้ใช้ต้องเห็น diff ชัดๆ ระหว่างของเดิมกับข้อเสนอ และต้อง revert ได้เสมอ — จึงวางฉบับแก้ไว้ **ติดกับคอลัมน์เดิม** และไฮไลต์ส่วนที่เปลี่ยน เพื่อให้อ่านเทียบซ้ายขวาได้โดยไม่ต้องสลับไปมา

## Workflow

```
0. เตรียมสคริปต์  →  1. รับ input (Sheet + Jira)  →  2. สรุป req + AC
→  2.5 Change Impact (เฉพาะ CR/ปรับ feature เดิม)  →  3. ตรวจเชิงกลไก
→  4. Semantic review ตาม checklist  →  5. เขียน TC ที่ขาด
→  6. สร้าง Review Sheet (สำเนาใหม่)  →  7. Upload + สรุปในแชท
```

ไม่ต้องหยุดถามระหว่างทาง ถ้า input ครบ (Sheet + ticket) — review เป็นงานที่ผู้ใช้อยากได้ผลทันที ข้อสงสัยด้าน requirement ให้รวบรวมใส่ใน summary เป็น "คำถามถึง BA/ลูกค้า" แทนการหยุดรอ หยุดถามเฉพาะเมื่อเปิดไฟล์ไม่ได้ หรือหา TC sheet ไม่เจอ

---

## Step 0 — เตรียมสคริปต์ (ทำครั้งเดียวต่อ session)

skill นี้มีสคริปต์สร้าง Review Sheet แนบอยู่ท้ายไฟล์ (**ภาคผนวก A**) ให้สร้างโฟลเดอร์ทำงานเฉพาะงานนี้ เช่น `review_<ชื่อไฟล์>_<timestamp>/` แล้วเขียนสคริปต์ลงไปด้วย heredoc ตามภาคผนวก A **ตรงตัว ห้ามแก้** (จะคัดลอกเองหรือดึงบล็อกนั้นจาก SKILL.md ไปรันก็ได้ ขอแค่เนื้อสคริปต์เหมือนเดิมทุกตัวอักษร) — ทำไม: สคริปต์คุมเรื่องที่พลาดง่าย (แทรกคอลัมน์ข้างเดิม, diff ตัวแดง, สีวนิลา, ล้าง Test Status, แปลง rich text ให้ Google Sheets อ่านได้) ถ้าเขียนเองใหม่ผลจะไม่เหมือนกันทุกครั้ง

ตรวจว่ามี `openpyxl` (`python3 -c "import openpyxl"`) ถ้าไม่มีให้ `pip install openpyxl --break-system-packages`

## Step 1 — รับ Input

**Google Sheet** (ลิงก์ `docs.google.com/spreadsheets/d/<fileId>/...`):
1. `get_file_metadata(fileId)` → ได้ชื่อไฟล์ + parent folder (ใช้วางไฟล์ review ไว้โฟลเดอร์เดียวกัน)
2. `download_file_content(fileId, exportMimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")` → decode base64 เซฟเป็น `original.xlsx` ใน working dir ของงานนี้
   - ใช้โฟลเดอร์ทำงานแยกเฉพาะงาน (เช่น `review_<ชื่อไฟล์>_<timestamp>/`) เก็บ original.xlsx, review.json, สคริปต์ — กันไฟล์ชนกันเมื่อมีหลายงานพร้อมกัน
   - การ export เป็น xlsx ทำให้ได้ style/width/tab ครบ ซึ่ง `read_file_content` (text) ไม่มี

**ไฟล์ .xlsx ที่แนบมา** → ใช้ตรงจาก uploads ได้เลย (copy มาทำงาน ไม่แก้ไฟล์ต้นฉบับ)

**Jira ticket** (Atlassian Rovo): `getAccessibleAtlassianResources` → cloudId → `getJiraIssue(key, expand="renderedFields,subtasks,issuelinks")` → อ่าน description, AC, comments (มัก clarify rule สำคัญ), subtasks
- ถ้าไม่มี ticket แต่ผู้ใช้วาง requirement มา → ใช้ข้อความนั้น
- ถ้าไม่มี requirement เลย → review คุณภาพ TC ได้ แต่แจ้งชัดใน summary ว่า **ไม่ได้ตรวจ Coverage (#1) และความถูกต้องเทียบ req** แล้วข้าม step 5

## Step 2 — สรุป Requirement เป็น AC list

แตก req เป็นรายข้อ `AC1, AC2, …` + business rules `BR1…` (ถ้า ticket ไม่ได้ number ไว้ ให้ตั้งเอง และบอกว่าตั้งเอง) — นี่คือแกนของ coverage map ทั้งหมด จดไว้ด้วยว่าแต่ละ AC อยู่ใน high-risk area ไหม (auth/authz, data loss, payment, PII, business rule ซ้อน, integration, high-traffic) เพราะ gap ใน high-risk area มีน้ำหนักมากกว่า

## Step 2.5 — Change Impact (เฉพาะงานที่แก้ feature เดิม)

**ขอบเขต:** ส่วนนี้มีไว้ตอบคำถามเดียว — *"TC ชุดที่ส่งมา review ครอบคลุมสิ่งที่เปลี่ยนและผลกระทบต่อของเดิมแล้วหรือยัง"* ไม่ใช่การ review code, ไม่หา bug ใน code, ไม่ออกแบบ test plan ใหม่ทั้งหมด ผลลัพธ์ทุกข้อต้องจบที่ TC: cover แล้ว (TC ไหน) / ยังไม่ cover (→ TC ใหม่) / TC เดิมต้องแก้หรือรันซ้ำ

**ทำเมื่อ:** ticket เป็น Change Request/Improvement, มี issue link ไปยัง ticket ของ feature เดิม (relates to, clones, is caused by, blocks), ผู้ใช้แนบ TC ชุดเดิม (regression) หรือ diff/PR, หรือขอให้ "เทียบกับของเดิม/ดู impact" ถ้าเป็น feature ใหม่ล้วนและไม่มีข้อมูลเหล่านี้ ข้าม step นี้

### A. Feature เดิม vs ใหม่

1. หา ticket เดิม: อ่าน issue links ของ ticket ใหม่ แล้ว `getJiraIssue` ทีละตัว (เฉพาะที่เกี่ยวกับ feature เดียวกัน) + อ่าน comments — feature อื่นที่พบว่ากระทบจาก code ถ้ารู้ ticket key (เช่น อ้างใน TC regression) ให้อ่านเพื่อดู rule เดิม แต่ไม่ต้องไล่ค้นหา ticket ที่ไม่มีใครอ้างถึง
2. หา TC ชุดเดิม (ถ้าผู้ใช้ให้ลิงก์/ไฟล์): download เป็น xlsx ไว้ที่ `regression.xlsx` — **อ่านอย่างเดียว ไม่แก้ ไม่ upload ทับ**
3. ทำตารางเทียบทีละ rule: ของเดิมว่าอย่างไร (อ้าง ticket/AC เดิม หรือ TC เดิม) → ของใหม่ว่าอย่างไร → ประเภท `เพิ่ม / เปลี่ยน / ลบ / คงเดิม` → TC ในชุดที่ review ตัวไหน cover
   - rule ที่ **เปลี่ยน** ต้องมี TC ทั้งฝั่งค่าใหม่และขอบเขตรอบค่าใหม่ (เช่น 5→3 ครั้ง ต้องมี 2 ครั้งไม่ล็อก + 3 ครั้งล็อก) และ TC ที่ยืนยันว่าค่าเดิมไม่ใช้แล้ว ถ้ามีความหมาย (เช่น ครั้งที่ 3 ต้องล็อกแล้ว ไม่ใช่รอถึง 5)
   - TC ในชุดใหม่ที่ยังใช้ค่าเดิม (เช่น ยังเขียน "15 นาที" ทั้งที่ใหม่เป็น 30) = finding #11 Correctness ระดับ Critical
   - rule ที่ **คงเดิม** แต่อยู่ใกล้จุดที่แก้ → ไม่ต้องเขียน TC ใหม่ ให้ชี้ TC เดิมที่ต้องรันซ้ำ

### B. จุดที่ code เปลี่ยน → เส้นทางที่กระทบ (ถ้ามีข้อมูล code)

แหล่งข้อมูล ตามลำดับ (ใช้เท่าที่มี ไม่ต้องขอทุกอย่าง):
- PR/commit ที่ผูกกับ ticket: ดูใน Development panel/remote links/comments ของ Jira; ถ้าเป็น GitHub และมี `gh` ใช้ได้ → `gh pr view <url> --json files,title,body` และ `gh pr diff <url>` หรือ `gh search prs <JIRA-KEY>`
- หาผู้เรียกใช้ (callers) ของไฟล์/ฟังก์ชันที่แก้: `gh search code "<symbol>" --repo <org/repo>` หรือ clone แบบตื้นแล้ว `grep -rn` — หรือใช้ผลค้นที่ dev/ผู้ใช้แนบมา
- ถ้าเข้าถึง code ไม่ได้ (Bitbucket/GitLab ที่ไม่มี connector, repo private ที่ไม่มีสิทธิ์) → ขอให้ผู้ใช้แนบ diff/PR หรือรายชื่อไฟล์ที่แก้ **ครั้งเดียว** ถ้าไม่มีก็ทำต่อจาก requirement อย่างเดียว และบอกชัดใน summary ว่า "ไม่ได้วิเคราะห์จาก code"

วิธีอ่าน diff สำหรับ QA (อ่านเพื่อหา *สิ่งที่ต้องเทส* ไม่ใช่ตรวจคุณภาพ code):
- ค่าคงที่/เงื่อนไขที่เปลี่ยน → boundary ใหม่ที่ต้องมี TC
- ข้อความ/ค่าที่แสดงผลที่เปลี่ยน → Expected ใน TC เดิมที่ต้องแก้
- **ไฟล์ shared/common/util ที่ถูกแก้** → ไล่ว่าใครใช้บ้าง ทุกหน้าที่ใช้คือ feature ที่อาจกระทบ แม้ ticket จะพูดถึงแค่หน้าเดียว — นี่คือจุดที่ regression หลุดบ่อยที่สุด
- ฟังก์ชันที่เปลี่ยน signature/ค่า default → caller เดิมที่ไม่ได้ส่งค่าใหม่ทำงานเหมือนเดิมไหม (เช่น token เพิ่ม purpose — flow เดิมยังใช้ได้ไหม)
- ไฟล์ใหม่/endpoint ใหม่ → flow ใหม่ที่ต้องมี TC ครบ happy + negative (หมดอายุ, ใช้ซ้ำ, สิทธิ์)
- side effect ใหม่ (audit log, email, notification) → ต้องมีอย่างน้อย 1 TC ยืนยัน

ทุกแถว impact ระบุ **ระดับความมั่นใจ**: `จาก code` (เห็นใน diff/ผลค้น caller) หรือ `อนุมานจาก req` — ทำไม: ผู้ใช้ต้องรู้ว่าข้อไหนมีหลักฐาน ข้อไหนควรถาม dev ยืนยัน

### C. สิ่งที่ได้จาก step นี้ (ไปรวมใน step 4-6)

- gap → TC ใหม่ใน `new_tcs` ใช้ `"check": 12` ต่อท้ายชีทที่ review เหมือน TC ใหม่อื่น (ไม่แยก tab) — TC ของ feature อื่นที่กระทบ ใช้ prefix module ของ feature นั้น (เช่น TC-REG-xxx) ต่อเลขจากชุด regression เดิมถ้ามี ถ้า feature นั้นยังไม่มี prefix ให้ตั้งใหม่ตามหลัก Test Case ID และบอกใน summary
- TC ในชุดที่ review ที่ขัดกับ behavior ใหม่ → finding ในชุดนั้นตามปกติ
- TC เดิม (regression) ที่ต้องทำอะไร → ใส่ในตาราง ไม่แก้ไฟล์เดิม: `ต้องแก้ Expected` (behavior เปลี่ยน — ใส่ข้อความ Expected ใหม่ให้เลย), `ต้องแก้ Test Data` (เช่น data เดิมไม่ผ่าน rule ใหม่ — ใส่ค่าใหม่ให้), `รันซ้ำ` (คงเดิมแต่อยู่ในเส้นที่กระทบ), `เลิกใช้` (behavior ถูกลบหรือมี TC ใหม่แทนแล้ว — บอกว่าแทนด้วย TC ไหน)
- feature ที่ตรวจแล้ว **ไม่กระทบ** → บันทึกพร้อมเหตุผล เพื่อให้ผู้ใช้เห็นว่าเช็คแล้ว ไม่ใช่ลืม

## Step 3 — ตรวจเชิงกลไก (Mechanical pass)

ไล่ทุกแถวหาปัญหาที่เช็คได้ตรงๆ ก่อน (อ่านไฟล์ด้วย openpyxl แล้ววนดูทีละแถว จะเร็วและไม่พลาดกว่าไล่ด้วยตา):
- ID ไม่ตรง `TC-[A-Z]+-\d{3}` หรือซ้ำ / Test Title ซ้ำกัน
- คอลัมน์ขาดจาก 12 คอลัมน์มาตรฐาน / ช่องบังคับว่าง (Module, Title, Steps, Expected)
- Expected มีคำกำกวม: ถูกต้อง, ปกติ, เหมาะสม, ตามที่ควรจะเป็น, ทำงานได้, OK, สำเร็จ (ลอยๆ — ไม่นับคำที่อยู่ในเครื่องหมายคำพูดเป็นข้อความบนจอ เช่น "บันทึกสำเร็จ")
- Steps/Expected ไม่เป็น numbered list, step ขึ้นต้นด้วย verb อังกฤษ (Click/Enter/Select/Navigate...), step มีหลาย action ("...แล้วคลิก"), expected ปนใน step ("ระบบควรแสดง...")
- Test Data เป็น placeholder: valid xxx, invalid xxx, test123, xxx, abc, ข้อมูลที่ถูกต้อง (ยกเว้นค่าที่ระบุความยาวไว้ชัด เช่น `"Abc" (3 ตัวอักษร)` ซึ่งเป็น BVA โดยตั้งใจ)
- Precondition ไม่ระบุบัญชี/role หรือหน้าเริ่มต้น
- emoji, prefix `[UI]`/`[DB]`/`[API]` ใน Expected
- Test Result Link / Linked Bug มีค่าค้างอยู่

ผลของ step นี้เป็นแค่วัตถุดิบ: **ตรวจทานทุกข้อก่อนใช้** และปัญหาเชิงความหมาย (ผิด req, ไม่ atomic, ทำซ้ำไม่ได้) ต้องอ่านเองใน step 4

## Step 4 — Semantic Review (ทีละ TC + ทั้งชุด)

อ่าน TC ทุกแถวจริงๆ อย่าดูแค่ผล step 3 ตรวจตาม checklist นี้ (ใช้หมายเลขเดียวกันในคอลัมน์ Review Checks เพื่อให้ filter ได้):

| # | Check | ถามตัวเองว่า |
|---|-------|-------------|
| 1 | Coverage | AC/BR ไหนไม่มี TC cover? TC ไหน cover AC ไม่ครบ (เช่น test แค่ min ไม่ test max)? |
| 2 | Risk coverage | High-risk area มี TC หนาแน่นกว่า low-risk ไหม? |
| 3 | Technique | ใช้ technique เหมาะไหม — field มี boundary แต่ไม่มี BVA ครบ 6 จุด (min-1, min, min+1, max-1, max, max+1) + empty? rule หลายเงื่อนไขแต่ไม่มี DT? มี state แต่ไม่ test invalid transition? |
| 4 | Negative/Edge ≥ Happy | นับจริง (จาก meta tag `[Type: ...]` ใน Comments/Notes หรืออ่านเนื้อ) ถ้าไม่ถึง → เขียน TC เพิ่ม |
| 5 | Atomic | 1 TC test 1 เรื่อง? ถ้า test หลาย rule ในตัวเดียว → แนะนำแตก (บอกว่าแตกเป็นอะไร) |
| 6 | Reproducible | คนอื่นทำตามได้โดยไม่ต้องถาม? Precondition มี role/บัญชี + หน้าเริ่มต้น + data ที่ต้องมีก่อน? Test Data เป็นค่าจริง? step ระบุ element ชัด? |
| 7 | Independent | TC พึ่งผลของ TC อื่นโดยไม่ระบุใน Precondition? |
| 8 | Format | 12 columns ครบ, ID `TC-[MODULE]-NNN` ไม่ซ้ำ, Steps/Expected เป็น numbered list, Test Result Link/Linked Bug เว้นว่าง |
| 9 | Browser scope | TC ที่ browser-specific ระบุใน Title/Notes? (default scope: Chrome/Safari/Firefox latest desktop) |
| 10 | ภาษา & Ambiguity | Expected ไม่มี "ถูกต้อง/ปกติ/เหมาะสม/ทำงานได้/OK/สำเร็จ(ลอยๆ)"; action verb ไทย (คลิกปุ่ม/กรอก/เลือก/เข้าสู่หน้า/ตรวจสอบ); ภาษาราชการ ไม่มี emoji/คำอวย/prefix `[UI]` |
| 11 | Correctness | Expected ขัดกับ req ไหม (เช่น req บอกล็อก 5 ครั้ง แต่ TC เขียน 3)? data ที่ใช้อยู่ใน partition ที่ Title อ้างจริงไหม (เช่น Title "ต่ำกว่า min" แต่ data ยาวเท่า min)? |
| 12 | Change Impact | (เฉพาะเมื่อทำ step 2.5) ทุก rule ที่เปลี่ยน/เพิ่มมี TC? feature อื่นที่ใช้ code ร่วมกันมี TC regression หรือ TC เดิมให้รันซ้ำ? |

**คอลัมน์ที่ไม่ review:**
- **Comments/Notes** — เป็นพื้นที่โน้ตของผู้เขียน (meta tags, dependency) ไม่ต้องเขียน finding หรือ fix ให้คอลัมน์นี้ (สคริปต์จะตัดทิ้งถ้าใส่มา) แต่ยังอ่านเพื่อเข้าใจบริบทได้ เช่น ดู `Covers ACn` ประกอบ coverage map
- **Test Status** — สคริปต์**ล้างเป็นค่าว่างให้อัตโนมัติ**ในฉบับ review (เป็นช่องกรอกตอน execute) ไม่มีคอลัมน์ (แก้ไข) และไม่ต้องเขียน finding

**ระดับชุด (ไม่ผูก TC เดียว):** TC ซ้ำ/รวมได้ (test สิ่งเดียวกัน data ต่างแต่ partition เดียวกัน), TC ที่ควรแตก, test data ไม่ realistic, non-functional ที่ req ไม่พูดถึงแต่ควรมี (security, session expired, double-submit, accessibility) — ให้ flag เป็น recommendation ใน summary

### Severity (ต่อ TC)

| Severity | เกณฑ์ | ทำไม |
|--------|-------|------|
| **Critical** | Expected ขัด req, TC execute ไม่ได้, TC test ผิด partition/ผิดเรื่อง | ถ้าส่งไป จะให้ผล pass/fail ผิด — อันตรายที่สุด |
| **Major** | Expected กำกวม, reproduce ไม่ได้ (data/precondition ขาด), ไม่ atomic, expected ปนใน step | tester แต่ละคนอาจตัดสินผลต่างกัน |
| **Minor** | ภาษา/format/meta tag/ID/verb อังกฤษ | ไม่กระทบผล test แต่ไม่พร้อมส่งลูกค้า |
| **Pass** | ไม่มีอะไรต้องแก้ | ใส่ไว้ให้เห็นว่าตรวจแล้ว ไม่ใช่ข้ามไป |

Severity ของ TC = ระดับที่หนักที่สุดของ finding ใน TC นั้น

### เขียน Findings + Fix

แต่ละ TC เขียนเป็น `findings` (1 finding = 1 ปัญหา) + `fix`:

```json
{"tc_id": "TC-LOGIN-003", "severity": "Critical",
 "findings": [
   {"check": 11, "column": "Test Steps", "summary": "จำนวนครั้งไม่ตรง AC3 (3 vs 5)",
    "issue": "Step 4 ทำซ้ำรวม 3 ครั้ง แต่ AC3 กำหนดล็อกเมื่อผิด 5 ครั้ง ถ้าระบบทำตาม req TC นี้จะ Fail ทั้งที่ระบบถูก"},
   {"check": 10, "column": "Expected Result", "summary": "มีเงื่อนไขที่ req ไม่ได้ระบุ",
    "issue": "ข้อ 4 \"ปุ่มถูก disable\" ไม่มีใน AC3 เป็น assumption ของผู้เขียน"}],
 "fix": {"Test Steps": "<ข้อความใหม่ทั้งช่อง>", "Expected Result": "<ข้อความใหม่ทั้งช่อง>"}}
```

- **check** = หมายเลข checklist (1-12) — สคริปต์แปลงเป็นชื่อเต็ม เช่น `[#11 Correctness - ตรงกับ requirement]`
- **column** = คอลัมน์ที่มีปัญหา (ชื่อตรงกับ header ในชีท)
- **summary** = สรุปปัญหาสั้นๆ ≤ 10 คำ ที่อ่านแล้วรู้เลยว่าผิดอะไร (ไม่ใช่ซ้ำชื่อ check) เช่น "ใช้คำกำกวม 'ถูกต้อง'", "Test Data เป็น placeholder", "step 2 มี 2 action"
- **issue** = อธิบายเต็ม: ตรงไหน + ผิดอย่างไร + ผลกระทบต่อการ test
- คอลัมน์ Review Checks จะแสดง `1. [#11 Correctness - ตรงกับ requirement] Test Steps: จำนวนครั้งไม่ตรง AC3 (3 vs 5)` และ Review Issue ข้อ 1 คู่กัน — ทำไม: ผู้อ่าน filter/scan ได้จาก Checks แล้วค่อยอ่านรายละเอียดใน Issue ข้อเดียวกัน
- **fix** = **ข้อความเต็มทั้งช่องที่แก้แล้ว** ของทุกคอลัมน์ที่ต้องแก้ เริ่มจากข้อความเดิมแล้วแก้เฉพาะจุด (อย่าเขียนใหม่ทั้งหมดถ้าไม่จำเป็น) — ทำไม: สคริปต์จะ diff กับของเดิมและไฮไลต์ส่วนที่เปลี่ยนเป็นตัวหนาสีแดง ถ้าเขียนใหม่ทั้งก้อน diff จะแดงทั้งช่องและเทียบไม่ได้
  - ต้องการให้ช่องว่าง → ใส่ `""` สคริปต์จะแสดง `[เว้นว่าง - ลบค่าเดิมออก]`
- ถ้าต้องแตก TC → finding บอกเหตุผล, fix ของ TC เดิมให้เหลือเรื่องเดียว แล้วเขียนส่วนที่แตกออกไว้ใน new_tcs
- ถ้าแก้ไม่ได้เพราะ req ไม่ระบุ (เช่น ไม่รู้ข้อความ error จริง) → เขียน fix ตาม assumption แล้วต่อท้าย `(TBC: รอยืนยันข้อความจริงจาก BA)` และใส่คำถามใน summary

กฎการเขียนเนื้อ fix และ TC ใหม่ = กฎเดียวกับ `qa-tc-full-workflow-standard` (สรุปไว้ในหัวข้อ **กฎการเขียน TC** ด้านล่าง — ทบทวนก่อนเขียน fix/TC ใหม่) เพราะ fix ที่ผิดกฎเองจะทำให้ reviewer หมดความน่าเชื่อถือ

## Step 5 — เขียน TC ที่ขาด

จาก gap ใน #1/#3/#4/#11/#12 เขียน TC ใหม่ครบ 12 คอลัมน์ ตามหัวข้อ **กฎการเขียน TC**:
- ID ต่อจากเลขสุดท้ายของ module นั้นในชีท (ไม่ใช้เลขซ้ำ) ในรูปแบบมาตรฐาน `TC-[MODULE]-NNN` 3 หลัก — ถ้าชีทเดิมใช้รูปแบบอื่น (เช่น `TC-REG-01`) ให้ flag ใน #8 พร้อม fix ID เป็น 3 หลัก แล้ว TC ใหม่ใช้ 3 หลักต่อจากนั้น (เช่น TC-REG-005) และบอกใน summary ว่าทั้งชุดควร rename
- Test Status / Test Result Link / Linked Bug / Owner เว้นว่าง
- Owner ที่มีค่าอยู่แล้วใน TC เดิม ไม่นับเป็น finding (ทีมอาจกำหนดผู้รับผิดชอบไว้แล้ว)
- Comments/Notes ของ TC ใหม่ยังเขียนตามรูปแบบทีมได้ (ไม่ได้ถูก review แต่เป็นส่วนหนึ่งของ TC):
  meta tags + `Covers ACn` + `[Review: NEW - <เหตุผลสั้น>]`
- `gap` = สรุปสั้น (ไปแสดงใน Review Checks) เช่น "AC6 ยังไม่มี TC", `reason` = เหตุผลเต็ม (ไปแสดงใน Review Issue), `check` = หมายเลข checklist ที่ทำให้เกิด gap (ปกติ 1, หรือ 3/4)
- ให้น้ำหนัก high-risk ก่อน ไม่ต้องเติมจนครบทุกความเป็นไปได้ถ้าไม่ช่วยลด risk — แต่ต้องทำให้ Negative+Edge ≥ Happy และ boundary ครบทุก field ที่มี limit

## Step 6 — สร้าง Review Sheet

เขียน `review.json` (schema ดูตัวอย่างใน Step 4 + Step 5 และโครงด้านล่าง) — ใส่ **ทุก TC** ใน `reviews` รวมตัวที่ Pass แล้วรัน:

```bash
python3 build_review_sheet.py original.xlsx review.json "[Review] <ชื่อเดิม> - YYYY-MM-DD.xlsx"
```

โครง `review.json`:

```json
{
  "sheet": "<ชื่อ tab ที่มี TC (ไม่ใส่ก็ได้ สคริปต์หาเองจาก header 'Test Case ID')>",
  "source_title": "<ชื่อไฟล์เดิม>", "jira": "SKL-482", "reviewer": "Claude", "date": "YYYY-MM-DD",
  "reviews": [ {"tc_id": "...", "severity": "Critical|Major|Minor|Pass", "findings": [...], "fix": {...}} ],
  "new_tcs": [ {"Test Case ID": "...", "Module / Feature": "...", "...": "ครบ 12 คอลัมน์",
                "check": 1, "gap": "AC6 ยังไม่มี TC", "reason": "..."} ],
  "summary_sections": [ {"title": "Coverage Map", "columns": ["AC", "..."], "rows": [["AC1", "..."]]} ],
  "extra_tabs": [ {"name": "Change Impact", "sections": [ {"title": "...", "note": "(ไม่ใส่ก็ได้)", "columns": [...], "rows": [[...]]} ]} ]
}
```

ถ้ามี TC ID ซ้ำในชีท ใส่ `"row": <เลขแถว>` ใน review นั้นเพื่อระบุแถว

ผลลัพธ์ (ชีท TC):

```
Test Case ID | Module | Test Title | Test Title (แก้ไข) | Precondition | ... | Expected Result | Expected Result (แก้ไข) | ... | Comments/Notes | Severity | Review Checks | Review Issue
```

- คอลัมน์ `(แก้ไข)` ถูกแทรก **ติดขวาคอลัมน์เดิม** เฉพาะคอลัมน์ที่มี fix อย่างน้อย 1 TC (ไม่สร้างคอลัมน์ว่าง) — cell ที่มี fix เป็นสีวนิลา ส่วนที่ต่างจากของเดิมเป็น**ตัวหนาสีแดง**
- เนื้อ TC เดิมคงค่าและ style เดิมทุก cell ยกเว้น Test Status ที่ถูกล้างเป็นค่าว่าง
- ไม่มีคอลัมน์ `Comments/Notes (แก้ไข)` และ `Test Status (แก้ไข)`
- Severity มีสีตามระดับ (Critical แดง / Major ส้ม / Minor ม่วงอ่อน / Pass เขียว)
- TC ใหม่ต่อท้ายชีท **ทั้งแถวสีวนิลา** Severity = `New TC`
- Freeze header + คอลัมน์ ID และเปิด filter ให้
- สคริปต์แปลงข้อความ rich text เป็น shared strings ให้อัตโนมัติ (ฟังก์ชัน `convert` ในสคริปต์) — ทำไม: ถ้าเก็บเป็น inline string, Google Sheets และ preview หลายตัวจะทิ้งสีแดง/ตัวหนา ทำให้มองไม่เห็นส่วนที่แก้ ถ้าแก้ไฟล์ด้วย openpyxl เองหลัง build ต้องแปลงซ้ำ (`python3 -c "from build_review_sheet import convert; convert('f.xlsx','f.xlsx')"`) ก่อน upload
- Tab `Review Summary`: จำนวนตาม Severity, วิธีอ่าน, checklist อ้างอิง + sections ที่ส่งมา

`summary_sections` ที่ต้องมี:
1. **Coverage Map** — `AC | รายละเอียด | Risk | TC เดิม | TC ใหม่ | สถานะ (Covered/Partial/GAP)`
2. **Checklist Result** — `# | Check | สถานะ (OK/WARN/FAIL) | หมายเหตุ`
3. **Type Breakdown** — `Type | ก่อน review | หลังเพิ่ม TC ใหม่` (ให้เห็นว่า Neg+Edge ≥ Happy หรือยัง — นับเฉพาะชุดที่ review ไม่รวม TC regression เดิม)
4. **ข้อเสนอระดับชุด** — `ประเภท (ซ้ำ/ควรแตก/Data/Non-functional) | TC ที่เกี่ยว | ข้อเสนอ`
5. **คำถามถึง BA/ลูกค้า** — `# | Quote req | ทำไมต้องถาม | ตัวเลือก A/B` (ถ้ามี)

ถ้าทำ step 2.5 ให้เพิ่ม `extra_tabs` ชื่อ **Change Impact** มี sections:
1. **แหล่งข้อมูลที่ใช้** — `แหล่ง | รายละเอียด (ticket key / PR / ไฟล์) | หมายเหตุ` รวมสิ่งที่**เข้าถึงไม่ได้**ด้วย
2. **Feature เดิม vs ใหม่** — `# | Rule | เดิม (อ้างอิง) | ใหม่ (อ้างอิง) | ประเภท | TC ที่ cover | สถานะ (Covered/Partial/GAP)`
3. **จุดที่ Code เปลี่ยน → ผลกระทบ** — `ไฟล์/ฟังก์ชัน | สิ่งที่เปลี่ยน | ใช้โดย (หน้าจอ/flow) | Feature ที่อาจกระทบ | ความมั่นใจ (จาก code/อนุมาน) | TC ที่ cover`
4. **TC เดิม (regression) ที่ต้องดำเนินการ** — `TC ID เดิม | Test Title | Action (ต้องแก้ Expected/ต้องแก้ Test Data/รันซ้ำ/เลิกใช้) | เหตุผล | ค่าใหม่ (Expected/Test Data ถ้าต้องแก้)`
5. **Feature ที่ตรวจแล้วไม่กระทบ** — `Feature | เหตุผลที่ไม่กระทบ | หลักฐาน`

ตรวจผลก่อนส่ง: `unmatched_ids` และ `unknown_fix_columns` ใน output ของสคริปต์ต้องว่าง (ชื่อคอลัมน์ใน fix ต้องตรงกับ header) และเปิดไฟล์อ่านกลับ (openpyxl) ว่าเซลล์ต้นฉบับไม่เปลี่ยน — ทำไม: คำมั่นหลักของ skill นี้คือไม่แตะของเดิม

## Step 7 — Upload + สรุป

**ถ้า input เป็น Google Sheet**: `create_file(title="[Review] <ชื่อเดิม> - YYYY-MM-DD", base64Content=<xlsx base64>, contentMimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", parentId=<folder เดิม>)` → Drive แปลงเป็น Google Sheet ให้ → ส่งลิงก์ ห้าม update/overwrite ไฟล์ต้นฉบับ
  - ได้ base64 ด้วย `base64 -w0 <file>.xlsx` แล้วคัดลอกทั้งก้อนลง `base64Content` ให้ตรงตัว (ไฟล์ review ปกติ ~20-40 KB) ถ้า Drive ตอบ error เรื่องแปลงไฟล์ แปลว่า base64 เพี้ยน ให้ส่งใหม่
  - ตัวหนาสีแดงในคอลัมน์ (แก้ไข) แสดงใน Google Sheets / Excel แต่ preview ไฟล์ในแชทไม่แสดงสีบางส่วนของข้อความ — ถ้าผู้ใช้บอกว่าไม่เห็นตัวแดงใน preview ให้แนะนำเปิดใน Google Sheets
**ถ้า input เป็น xlsx / Drive ใช้ไม่ได้**: เซฟไฟล์ไว้ในโฟลเดอร์ output ของ session (เช่น `/mnt/user-data/outputs/` ถ้ามี หรือ path ที่ผู้ใช้กำหนด) แล้วส่งไฟล์ให้ผู้ใช้

สรุปในแชท (กระชับ ไม่ recap ทุก TC — รายละเอียดอยู่ในชีทแล้ว):

```markdown
## ผล Review — <ชื่อชีท> (<Jira key>)
ลิงก์: <Review Sheet>

- TC เดิม N ตัว (Severity): Critical a | Major b | Minor c | Pass d
- TC ใหม่ที่เสนอ: m ตัว (cover AC…)
- Coverage: x/y AC covered — GAP: AC3 (lockout), AC5 (…)
- Neg+Edge : Happy = ก่อน p:q → หลัง p':q'
- (ถ้าทำ Change Impact) rule ที่เปลี่ยน x ข้อ — GAP y | feature อื่นที่กระทบ: … | TC เดิมต้องแก้ a ตัว, รันซ้ำ b ตัว

### ต้องแก้ก่อนส่งลูกค้า (Critical/Major เด่นๆ)
| TC ID | Issue | Suggested Fix (ย่อ) |
|-------|-------|---------------------|

### คำถามถึง BA/ลูกค้า
1. …
```

ตาราง "ต้องแก้" ใส่เฉพาะ Critical + Major ที่สำคัญสุด ≤ 10 แถว เพื่อให้ผู้ใช้เห็น priority ทันที ส่วนที่เหลือดูในชีท

---

## หลักการ Reviewer ที่ดี

- **เจาะจง ไม่ลอย**: "Expected ข้อ 2 ไม่บอกข้อความ error" ดีกว่า "Expected ยังไม่ชัด"
- **ทุก issue ต้องมี fix ที่ใช้ได้จริง** — ถ้าเสนอ fix ไม่ได้ แปลว่ายังไม่เข้าใจปัญหาพอ หรือ req ขาด (→ ไปเป็นคำถาม)
- **อย่าจับผิดจนเกินเหตุ**: ถ้า TC ดีแล้วให้ Pass ไม่ต้องหา Minor มาใส่ให้ครบ — noise ทำให้ issue จริงจมหาย
- **ไม่เปลี่ยน scope**: ถ้าเห็นว่า req เองมีปัญหา (ขัดกันเอง, ขาด non-functional) ให้ flag ใน summary ไม่ใช่ลงโทษ TC
- **Assumption ชัด**: ทุกครั้งที่ตีความ req เอง ให้บอกใน summary

---

## กฎการเขียน TC (ใช้กับข้อความใน (แก้ไข) และ TC ใหม่)

ย่อจาก `qa-tc-full-workflow-standard` เพื่อให้ fix/TC ใหม่ สอดคล้องกับ TC ที่ทีมเขียนอยู่แล้ว

### 12 Columns

`Test Case ID | Module / Feature | Test Title | Precondition | Test Data | Test Steps | Expected Result | Test Status | Test Result Link | Linked Bug (if any) | Owner | Comments/Notes`

- Test Status, Test Result Link, Linked Bug, Owner → เว้นว่าง (กรอกตอน execute)
- Type ไม่ใช่ column แยก — เป็น meta tag ใน Comments/Notes

### ภาษา

1. ภาษาเขียนทางการ/ราชการ กระชับ ไม่ใช้ภาษาพูด
2. Action verb ใน Test Steps เป็น**ภาษาไทย**: คลิกปุ่ม, กรอก, เลือก, เข้าสู่หน้า, ตรวจสอบ, เปิด, ปิด, อัปโหลด, เลื่อน, รอ (ไม่ใช้ Click/Enter/Navigate)
3. ศัพท์เทคนิคที่ไม่มีคำไทยชัด (URL, API, token, Dashboard, checkbox) คงอังกฤษได้ แต่กริยา/คำเชื่อมเป็นไทย เช่น "ระบบพาไปยังหน้า Dashboard"
4. ห้ามคำกำกวมใน Expected: ถูกต้อง, ปกติ, เหมาะสม, ตามที่ควรจะเป็น, ทำงานได้, OK, สำเร็จ (ลอยๆ)
5. ห้าม emoji, คำอวยแบบ AI ("ราบรื่น", "ประสบการณ์ที่ดี"), prefix `[UI]` `[DB]` `[API]`
6. ใช้ "ผู้ใช้งาน", "เข้าสู่ระบบ", "บันทึกข้อมูล" ในประโยคบรรยาย

### Test Case ID
`TC-[MODULE ย่อ]-[NNN]` อักษรใหญ่ 3-6 ตัว หนึ่ง module ใช้ prefix เดียว เช่น LOGIN, REG, PWD, PROFILE, USER, PERM, SEARCH, CART, PAY, ORDER, DASH, UPLOAD, NOTI — ถ้าชีทมี convention อยู่แล้ว ยึดของชีท

### Test Title
`[สิ่งที่ทดสอบ] + [เงื่อนไข/ข้อมูล/สถานะ] + [role/browser ถ้าจำเป็น]`
- ดี: `เข้าสู่ระบบด้วยรหัสผ่านที่ไม่ถูกต้อง ครบ 5 ครั้ง แล้วระบบล็อกบัญชี`
- ไม่ดี: `ทดสอบ Login`, `Test case 1`, `ตรวจสอบว่าระบบทำงานถูกต้อง`

### Precondition (numbered list)
1. บัญชี/role ที่ใช้ (เช่น "เข้าสู่ระบบด้วยบัญชีสิทธิ์ Admin: qa.admin+001@example.com" หรือ "ยังไม่เข้าสู่ระบบ")
2. หน้าจอเริ่มต้น
3. ข้อมูลที่ต้องมีอยู่ก่อน
4. สถานะของ object (ถ้า state-dependent)
5. browser/device/feature flag (ถ้าเจาะจง)

### Test Data — ค่าจริงเสมอ
Default: Email `test+001@example.com`…; Phone `0900000001`–`0900000099`; ชื่อไทย `ทดสอบ ระบบ001`; ชื่ออังกฤษ `Test User001`; Password valid `Test@1234`
ห้าม: "valid email", "test123", "xxx", "ข้อมูลที่ถูกต้อง"
BVA ต้องระบุค่าจริงพร้อมความยาว เช่น `Password: "Test@12" (7 ตัวอักษร)`
- ค่ายาวมาก (เช่น ชื่อ 100/101 ตัวอักษร) เขียนเป็นรูปแบบที่สร้างซ้ำได้ เช่น `"ทดสอบ ระบบ" + "ก" x 90 (รวม 100 ตัวอักษร)` เพื่อให้ tester สร้างค่าเดิมได้ทุกครั้ง
- field ที่ความยาวตายตัว (เช่น เบอร์โทร 10 หลัก, min = max) ทดสอบ N-1, N, N+1 และค่าว่าง

### Test Steps
```
1. เข้าสู่หน้า /login
2. กรอก "test+001@example.com" ในช่อง Email
3. กรอก "Test@1234" ในช่อง Password
4. คลิกปุ่ม "เข้าสู่ระบบ"
```
1 step = 1 action, ระบุ element ชัด, ไม่ใส่ expected ใน step

### Expected Result (numbered list map กับ step ที่มีผลสังเกตได้)
```
1. ระบบแสดงหน้า "เข้าสู่ระบบ" พร้อมช่อง Email และ Password
4. ระบบพาไปยังหน้า Dashboard และแสดงชื่อ "Test User001" ที่มุมขวาบน
```
ระบุสิ่งที่เห็นได้จริง: ข้อความ (quote ตรงตัว), URL ปลายทาง, สถานะปุ่ม, ค่าที่แสดง, ค่าใน DB, API response
1 ข้อ = 1 ผลลัพธ์ ไม่ยัดหลายผลในข้อเดียว

### Comments/Notes
`[Priority: P1] [Type: Negative] [Technique: BVA] Covers AC2. <notes อื่น>`
- Priority: P1–P4 | Type: Functional / UI / Negative / Boundary / Security / Compatibility / Usability / State | Technique: EP / BVA / DT / ST / UC
- Comments/Notes ไม่ถูก review แต่ TC ใหม่ยังเขียนตามรูปแบบนี้
- TC ใหม่จาก review เพิ่ม `[Review: NEW - <เหตุผล>]`
- TC ที่ควร automate: `[Automation: Recommended - <เหตุผล>]`

---

## ภาคผนวก A — สคริปต์ build_review_sheet.py

รันคำสั่งนี้ในโฟลเดอร์ทำงาน (คัดลอกทั้งก้อน):

```bash
cat > build_review_sheet.py <<'PYEOF'
#!/usr/bin/env python3
"""Build the review copy of a TC workbook: python build_review_sheet.py original.xlsx review.json output.xlsx"""
import difflib
import zipfile
import os
import json
import re
import sys
from collections import Counter
from copy import copy

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VANILLA = PatternFill("solid", fgColor="FFF3E5AB")      # suggestion / new content (vanilla)
FIX_HDR = PatternFill("solid", fgColor="FFE0C068")      # header of "(แก้ไข)" columns (darker vanilla)
REVIEW_HDR = PatternFill("solid", fgColor="FF9FC5E8")   # header of review columns
SUM_HDR = PatternFill("solid", fgColor="FFDAEEF3")
SEV_FILL = {
    "Critical": PatternFill("solid", fgColor="FFF4CCCC"),
    "Major": PatternFill("solid", fgColor="FFFCE5CD"),
    "Minor": PatternFill("solid", fgColor="FFD9D2E9"),   # lavender, so it does not clash with vanilla
    "Pass": PatternFill("solid", fgColor="FFD9EAD3"),
    "New TC": VANILLA,
}
SEV_ORDER = ["Critical", "Major", "Minor", "Pass", "New TC"]
WRAP = Alignment(wrap_text=True, vertical="top")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="FFBFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_NORMAL = InlineFont(rFont="Calibri", sz=11)
F_CHANGED = InlineFont(rFont="Calibri", sz=11, b=True, color="FFC00000")
FIX_SUFFIX = " (แก้ไข)"
# columns that are not reviewed: no findings, no "(แก้ไข)" column
NOT_REVIEWED = {"comments/notes", "test status"}
# columns cleared to empty in the review copy (execution-time fields must be blank before delivery)
CLEAR_COLS = {"test status"}

CHECKS = {
    1: "Coverage - ครอบคลุม AC/BR",
    2: "Risk coverage - ความหนาแน่นใน high-risk area",
    3: "Technique - เลือกใช้เทคนิค (EP/BVA/DT/ST) ไม่ครบ",
    4: "Negative/Edge >= Happy",
    5: "Atomic - 1 TC ทดสอบ 1 เรื่อง",
    6: "Reproducible - ทำซ้ำได้โดยไม่ต้องถาม",
    7: "Independent - ไม่พึ่งผลของ TC อื่น",
    8: "Format - คอลัมน์/ID/numbered list",
    9: "Browser scope",
    10: "ภาษา & Ambiguity",
    11: "Correctness - ตรงกับ requirement",
    12: "Change Impact - ครอบคลุมจุดที่เปลี่ยนและ feature เดิมที่ได้รับผลกระทบ",
}


def norm(s):
    return " ".join(str(s or "").split()).lower()


def header_row(ws):
    for r in range(1, min(ws.max_row, 15) + 1):
        if any(norm(c.value) == "test case id" for c in ws[r]):
            return r
    return None


def find_sheet(wb, name):
    if name and name in wb.sheetnames:
        return wb[name], header_row(wb[name])
    for ws in wb.worksheets:
        hr = header_row(ws)
        if hr:
            return ws, hr
    sys.exit("ไม่พบ header 'Test Case ID'")


def check_label(c):
    m = re.match(r"#?\s*(\d+)", str(c))
    if m and int(m.group(1)) in CHECKS:
        return f"#{m.group(1)} {CHECKS[int(m.group(1))]}"
    return str(c)


def diff_rich(old, new):
    """Rich text of `new` with parts that differ from `old` in bold red."""
    old, new = str(old or ""), str(new or "")
    if not new.strip():
        return CellRichText([TextBlock(F_CHANGED, "[เว้นว่าง - ลบค่าเดิมออก]")]) if old.strip() else None
    if not old.strip():
        return CellRichText([TextBlock(F_CHANGED, new)])
    # tokens: ASCII words/numbers stay whole (so "001"->"005" highlights the number), Thai/other chars one by one
    tok = lambda s: re.findall(r"[A-Za-z0-9_@.]+|\s+|.", s, re.S)
    a, b = tok(old), tok(new)
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    segs = []  # (changed?, text)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if j2 > j1:
            segs.append((op != "equal", "".join(b[j1:j2])))
    # absorb tiny unchanged islands (<3 chars, not whitespace-only) into changes to avoid noisy speckles
    merged = []
    for k, (chg, txt) in enumerate(segs):
        if not chg and len(txt) < 3 and txt.strip() and 0 < k < len(segs) - 1:
            chg = True
        if merged and merged[-1][0] == chg:
            merged[-1] = (chg, merged[-1][1] + txt)
        else:
            merged.append((chg, txt))
    if not any(c for c, _ in merged):
        return new
    return CellRichText([TextBlock(F_CHANGED if c else F_NORMAL, t) for c, t in merged if t])


def numbered(items):
    return "\n".join(f"{i}. {t}" for i, t in enumerate(items, 1))


def normalize_review(rv):
    sev = rv.get("severity") or rv.get("result") or "Pass"
    if sev.upper() == "NEW":
        sev = "New TC"
    findings = rv.get("findings")
    if findings is None:  # legacy
        checks, issues = rv.get("checks") or [], rv.get("issues") or []
        findings = [{"check": checks[i] if i < len(checks) else "", "column": "", "issue": iss}
                    for i, iss in enumerate(issues)]
    return sev, findings, rv.get("fix") or {}


def copy_style(src, dst):
    if src.has_style:
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.number_format, dst.protection = copy(src.alignment), src.number_format, copy(src.protection)


# --- rich text must live in sharedStrings, otherwise Google Sheets drops the red runs ---
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SST_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
SST_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

CELL_RE = re.compile(r'<c ([^>]*?)t="inlineStr"([^>]*)><is>(.*?)</is></c>', re.S)


def convert(src, dst):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    existing = files.get("xl/sharedStrings.xml")
    items, index = [], {}
    if existing:
        body = existing.decode("utf-8")
        items = re.findall(r"<si>(.*?)</si>", body, re.S)
        index = {s: i for i, s in enumerate(items)}

    total = 0
    for name in [n for n in files if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]:
        xml = files[name].decode("utf-8")

        def repl(m):
            nonlocal total
            inner = m.group(3)
            if inner not in index:
                index[inner] = len(items)
                items.append(inner)
            total += 1
            return f'<c {m.group(1)}t="s"{m.group(2)}><v>{index[inner]}</v></c>'

        xml = CELL_RE.sub(repl, xml)
        # empty cells written as inlineStr without <is> are invalid in strict readers: make them plain styled cells
        xml = re.sub(r'(<c [^>]*?) t="inlineStr"([^>]*)></c>', r"\1\2/>", xml)
        xml = re.sub(r'(<c [^>]*?) t="inlineStr"([^>]*)/>', r"\1\2/>", xml)
        files[name] = xml.encode("utf-8")

    sst = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<sst xmlns="{NS_MAIN}" count="{total}" uniqueCount="{len(items)}">'
           + "".join(f"<si>{s}</si>" for s in items) + "</sst>")
    files["xl/sharedStrings.xml"] = sst.encode("utf-8")

    ct = files["[Content_Types].xml"].decode("utf-8")
    if "sharedStrings.xml" not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="/xl/sharedStrings.xml" ContentType="{SST_CT}"/></Types>')
        files["[Content_Types].xml"] = ct.encode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "sharedStrings" not in rels:
        ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)] or [0]
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="rId{max(ids) + 1}" Type="{SST_REL}" Target="sharedStrings.xml"/></Relationships>')
        files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        order = ["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]
        for n in order:
            z.writestr(n, files[n])
    return total, len(items)


def write_sections(ws, sections, row=0):
    """Write titled tables (title / columns / rows) one after another; returns last row used."""
    for sec in sections:
        row += 2 if row else 1
        ws.cell(row=row, column=1, value=sec["title"]).font = Font(bold=True, size=12)
        if sec.get("note"):
            row += 1
            ws.cell(row=row, column=1, value=sec["note"]).alignment = Alignment(wrap_text=False)
        row += 1
        for j, h in enumerate(sec.get("columns", []), 1):
            c = ws.cell(row=row, column=j, value=h)
            c.font, c.fill, c.border, c.alignment = BOLD, SUM_HDR, BOX, WRAP
        for data in sec.get("rows", []):
            row += 1
            for j, v in enumerate(data, 1):
                c = ws.cell(row=row, column=j, value=v)
                c.border, c.alignment = BOX, WRAP
                if str(v).strip().upper() in ("GAP", "NEW", "NEW TC", "NOT COVERED"):
                    c.fill = VANILLA
    return row


def main(src_path, review_path, out):
    review = json.load(open(review_path, encoding="utf-8"))
    wb = load_workbook(src_path)
    ws, hr = find_sheet(wb, review.get("sheet"))
    headers = {c.column: str(c.value).strip() for c in ws[hr] if c.value is not None}
    by_name = {norm(v): k for k, v in headers.items()}
    id_col = by_name["test case id"]
    last_col = max(headers)

    last_row = hr
    for r in range(hr + 1, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in headers):
            last_row = r
    rows_by_id = {}
    for r in range(hr + 1, last_row + 1):
        rows_by_id.setdefault(str(ws.cell(row=r, column=id_col).value or "").strip(), []).append(r)

    # resolve reviews to rows
    resolved, unmatched, unknown_cols = {}, [], set()
    for rv in review.get("reviews", []):
        r = rv.get("row") or (rows_by_id.get(str(rv["tc_id"]).strip()) or [None])[0]
        if not r:
            unmatched.append(rv["tc_id"])
            continue
        sev, findings, fix = normalize_review(rv)
        findings = [f for f in findings if norm(f.get("column")) not in NOT_REVIEWED]
        if not findings and sev not in ("Pass", "New TC"):
            sev = "Pass"
        fix_cols = {}
        for k, v in fix.items():
            if norm(k) in NOT_REVIEWED:
                continue
            c = by_name.get(norm(k))
            if c:
                fix_cols[c] = v
            else:
                unknown_cols.add(k)
        resolved[r] = (sev, findings, fix_cols)
    fixed_cols = sorted({c for _, _, f in resolved.values() for c in f})

    # build new column order
    layout = []  # (kind, orig_col)
    for c in range(1, last_col + 1):
        layout.append(("orig", c))
        if c in fixed_cols:
            layout.append(("fix", c))
    review_cols = [("Severity", 12), ("Review Checks", 34), ("Review Issue", 52)]

    title, idx = ws.title, wb.sheetnames.index(ws.title)
    nws = wb.create_sheet(title + "__tmp", idx)
    nws.sheet_properties = copy(ws.sheet_properties)

    # rows above + header
    for r in range(1, last_row + 1):
        if ws.row_dimensions[r].height:
            nws.row_dimensions[r].height = ws.row_dimensions[r].height
        for j, (kind, c) in enumerate(layout, 1):
            src = ws.cell(row=r, column=c)
            dst = nws.cell(row=r, column=j)
            if kind == "orig":
                dst.value = None if (r > hr and norm(headers[c]) in CLEAR_COLS) else src.value
                copy_style(src, dst)
            elif r == hr:
                dst.value = headers[c] + FIX_SUFFIX
                dst.fill, dst.font, dst.alignment, dst.border = FIX_HDR, BOLD, WRAP, BOX
            elif r in resolved and c in resolved[r][2]:
                dst.value = diff_rich(src.value, resolved[r][2][c])
                dst.fill, dst.alignment, dst.border = VANILLA, WRAP, BOX
            elif r > hr:
                dst.border = BOX
    for j, (kind, c) in enumerate(layout, 1):
        w = ws.column_dimensions[get_column_letter(c)].width or 20
        nws.column_dimensions[get_column_letter(j)].width = max(w, 30) if kind == "fix" else w

    start = len(layout) + 1
    for i, (name, width) in enumerate(review_cols):
        cell = nws.cell(row=hr, column=start + i, value=name)
        cell.fill, cell.font, cell.alignment, cell.border = REVIEW_HDR, BOLD, WRAP, BOX
        nws.column_dimensions[get_column_letter(start + i)].width = width

    counts = Counter()
    for r in range(hr + 1, last_row + 1):
        if r not in resolved:
            continue
        sev, findings, fix_cols = resolved[r]
        counts[sev] += 1
        def check_line(f):
            line = f"[{check_label(f.get('check', ''))}]"
            if f.get("column"):
                line += f" {f['column']}"
            if f.get("summary"):
                line += f": {f['summary']}"
            return line
        checks = numbered([check_line(f) for f in findings])
        issues = numbered([f.get("issue", "") for f in findings])
        for i, v in enumerate([sev, checks or None, issues or None]):
            cell = nws.cell(row=r, column=start + i, value=v)
            cell.alignment, cell.border = WRAP, BOX
        nws.cell(row=r, column=start).fill = SEV_FILL.get(sev, VANILLA)

    # new TCs
    r = last_row
    for tc in review.get("new_tcs", []):
        r += 1
        counts["New TC"] += 1
        lookup = {norm(k): v for k, v in tc.items()}
        for j, (kind, c) in enumerate(layout, 1):
            cell = nws.cell(row=r, column=j, value=lookup.get(norm(headers[c])) if kind == "orig" else None)
            cell.fill, cell.alignment, cell.border = VANILLA, WRAP, BOX
        vals = ["New TC", f"1. [{check_label(tc.get('check', 1))}] {tc.get('gap', 'เพิ่ม TC ที่ขาด')}",
                f"1. {tc.get('reason', '')}"]
        for i, v in enumerate(vals):
            cell = nws.cell(row=r, column=start + i, value=v)
            cell.fill, cell.alignment, cell.border = VANILLA, WRAP, BOX

    nws.freeze_panes = nws.cell(row=hr + 1, column=2)
    nws.auto_filter.ref = f"A{hr}:{get_column_letter(start + len(review_cols) - 1)}{r}"
    wb.remove(ws)
    nws.title = title

    # summary tab
    if "Review Summary" in wb.sheetnames:
        del wb["Review Summary"]
    sm = wb.create_sheet("Review Summary")
    sm.column_dimensions["A"].width = 26
    for col in "BCDEFG":
        sm.column_dimensions[col].width = 34
    row = 1
    sm.cell(row=row, column=1, value="TC Review Summary").font = Font(bold=True, size=14)
    for label, key in [("ไฟล์ต้นฉบับ", "source_title"), ("Jira", "jira"), ("Reviewer", "reviewer"), ("วันที่", "date")]:
        if review.get(key):
            row += 1
            sm.cell(row=row, column=1, value=label).font = BOLD
            sm.cell(row=row, column=2, value=review[key])
    row += 2
    for j, h in enumerate(["Severity", "จำนวน"], 1):
        c = sm.cell(row=row, column=j, value=h)
        c.font, c.fill = BOLD, SUM_HDR
    for k in SEV_ORDER:
        row += 1
        sm.cell(row=row, column=1, value=k).fill = SEV_FILL[k]
        sm.cell(row=row, column=2, value=counts.get(k, 0))
    row += 2
    sm.cell(row=row, column=1, value="วิธีอ่าน").font = BOLD
    legend = [
        (VANILLA, "คอลัมน์ \"(แก้ไข)\" อยู่ติดกับคอลัมน์เดิม = ข้อความฉบับแก้ทั้งช่อง พร้อม copy ไปแทน"),
        (None, "ไม่ review คอลัมน์ Comments/Notes และ Test Status (Test Status ถูกล้างเป็นค่าว่างในฉบับนี้)"),
        (None, "ตัวอักษรหนาสีแดงในคอลัมน์ (แก้ไข) = ส่วนที่ต่างจากของเดิม"),
        (VANILLA, "แถวสีวนิลาทั้งแถว = TC ใหม่ที่เสนอเพิ่ม (Severity = New TC)"),
        (None, "Review Checks ข้อที่ n คู่กับ Review Issue ข้อที่ n"),
    ]
    for fill, text in legend:
        row += 1
        c = sm.cell(row=row, column=2, value=text)
        if fill:
            c.fill = fill
    row += 1
    sm.cell(row=row, column=1, value="Checklist อ้างอิง").font = BOLD
    for n, name in CHECKS.items():
        row += 1
        sm.cell(row=row, column=2, value=f"#{n} {name}")

    write_sections(sm, review.get("summary_sections", []), row)

    # extra tabs, e.g. "Change Impact" (feature diff, code impact, regression TCs to update/re-run)
    for tab in review.get("extra_tabs", []):
        name = tab["name"][:31]
        if name in wb.sheetnames:
            del wb[name]
        xt = wb.create_sheet(name)
        xt.column_dimensions["A"].width = 22
        for col in "BCDEFGH":
            xt.column_dimensions[col].width = 34
        write_sections(xt, tab.get("sections", []))

    wb.save(out)
    # rich-text runs must live in sharedStrings, otherwise Google Sheets / previews drop the red highlight
    convert(out, out)
    print(json.dumps({"output": out, "sheet": title, "counts": dict(counts),
                      "fix_columns": [headers[c] + FIX_SUFFIX for c in fixed_cols],
                      "unmatched_ids": unmatched, "unknown_fix_columns": sorted(unknown_cols)},
                     ensure_ascii=False))


if __name__ == "__main__":
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    main(*sys.argv[1:])
PYEOF
```