# qa-tc-review

Claude Skill สำหรับ **Review Test Case** ที่เขียนไว้แล้ว (Google Sheet / .xlsx) เทียบกับ Jira ticket ตาม QA checklist
ใช้เพื่อ review เท่านั้น ไม่ได้เขียน TC ใหม่ทั้งชุด

## Skill ทำอะไรได้

- สร้าง **Sheet สำเนาใหม่** โดยไม่แตะต้นฉบับ
- เพิ่มคอลัมน์ **(แก้ไข)** ข้างคอลัมน์เดิม ไฮไลต์สีวนิลา และแสดงส่วนที่แก้เป็น **ตัวแดงตัวหนา** (แสดงผลบน Google Sheets)
- เพิ่มคอลัมน์ **Severity** (Critical / Major / Minor / Pass / New TC), **Review Checks** และ **Review Issue**
- เพิ่ม **TC ที่ขาด** เป็นแถวใหม่สีวนิลา
- เพิ่มแท็บ **Review Summary**
- **Change Impact** (ทำอัตโนมัติเมื่อเป็น CR หรือ ticket มี linked ticket ของ feature เดิม): เทียบ feature เดิมกับใหม่ วิเคราะห์ PR diff ว่ากระทบส่วนไหน และระบุ regression TC ที่ต้องแก้หรือต้องรันซ้ำ
- ไม่ review Comments/Notes และล้างค่า Test Status ให้ว่าง

Checklist ที่ใช้ review: Coverage, Risk, Technique, Negative/Edge ≥ Happy, Atomic, Reproducible, Independent, Format, Browser scope, ภาษา & Ambiguity, Correctness, Change Impact

## ติดตั้ง

**วิธีที่ 1 (ง่ายสุด):** ดาวน์โหลด [`dist/qa-tc-review.skill`](dist/qa-tc-review.skill)
แล้วใน Claude ไปที่ **Settings → Capabilities → Skills → Upload skill** แล้วเลือกไฟล์นี้

**วิธีที่ 2 (Claude Code):** คัดลอกโฟลเดอร์ `qa-tc-review/` ไปไว้ที่ `~/.claude/skills/`

## Connector ที่แนะนำ

| Connector | ใช้ทำอะไร | ถ้าไม่มี |
|---|---|---|
| Google Drive | ดึง Sheet ต้นฉบับและอัปโหลด Sheet ผลลัพธ์ | แนบไฟล์ .xlsx มาแทน แล้วรับไฟล์ผลลัพธ์เป็น .xlsx |
| Atlassian (Jira) | อ่าน ticket และ linked ticket | วางเนื้อหา requirement ในแชทแทน |
| GitHub (`gh` CLI) | อ่าน PR diff ตอนทำ Change Impact | Skill จะขอ diff หนึ่งครั้ง หรือวิเคราะห์จาก requirement แทน (ติดป้าย "อนุมานจาก req") |

## วิธีใช้

```
Review TC ใน Sheet นี้ เทียบกับ SKL-123
<ลิงก์ Google Sheet>
```

ตัวอย่างงาน CR:

```
Review TC ของ SKL-530 (CR ของ SKL-482) ใน Sheet นี้ เช็ค regression ด้วย
PR: https://github.com/<org>/<repo>/pull/1234
```

## โครงสร้าง repo

```
qa-tc-review/SKILL.md     ตัว Skill (script สร้าง Sheet ฝังอยู่ใน Appendix A)
dist/qa-tc-review.skill   ไฟล์ packaged พร้อมอัปโหลดเข้า Claude
```

## การอัปเดต Skill

แก้ `qa-tc-review/SKILL.md` แล้ว package ใหม่ด้วยคำสั่งนี้ (ไฟล์ .skill คือ zip ของโฟลเดอร์):

```bash
cd <repo> && rm -f dist/qa-tc-review.skill && zip -r dist/qa-tc-review.skill qa-tc-review
```

## Changelog

- **v2**: เพิ่ม Change Impact (checklist #12) และแท็บ Change Impact
- **v1**: Review TC + คอลัมน์ (แก้ไข) + Severity + TC ที่ขาด + Review Summary
