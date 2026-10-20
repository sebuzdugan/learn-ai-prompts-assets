# Invitation documents: verification report

Prepared for Sebastian-Eugen Buzdugan and Codruta Costea. Sources were checked on **20 October 2026**.

## Bottom line

**None of the three organizers publishes an invitation-letter template, attendee confirmation, certificate or sample ticket.** As a result, all six documents are **independently designed drafts**. Each one is a *request for an invitation letter* written in the attendee's own voice. None reproduces an official format, and none is presented as issued by an organizer.

| Event | Official template found? | What was produced |
|---|---|---|
| AI Engineers Meetup Bangkok | No | A request for attendance confirmation, addressed to the Meetup host |
| AI for Science Congress 2026 | No | A request for attendee registration and an invitation letter, addressed to the organizing committee |
| MAIXPO 2026 | No. The site gives a procedure (an email with listed details) but no template | A request for a visa invitation letter that follows MAIXPO's published procedure |

None of the documents contains organizer signatures, stamps, seals, registration or ticket numbers, QR codes, payment confirmations, or any claim that registration is complete. Fields that only an organizer can supply (registration ID, ticket reference) are left blank. Status rows are checkboxes for the attendee to tick. Organizer logos are left out on purpose, because using them on a document the organizer did not issue would imply endorsement. Each design takes its palette and typefaces from the organizer's own website instead.

## Files

| | Sebastian-Eugen Buzdugan | Codruta Costea |
|---|---|---|
| Meetup, Bangkok | [PDF](pdf/01_AI-Engineers-Meetup-Bangkok_Sebastian-Eugen-Buzdugan.pdf) · [DOCX](docx/01_AI-Engineers-Meetup-Bangkok_Sebastian-Eugen-Buzdugan.docx) | [PDF](pdf/01_AI-Engineers-Meetup-Bangkok_Codruta-Costea.pdf) · [DOCX](docx/01_AI-Engineers-Meetup-Bangkok_Codruta-Costea.docx) |
| AI for Science Congress | [PDF](pdf/02_AI-for-Science-Congress-2026_Sebastian-Eugen-Buzdugan.pdf) · [DOCX](docx/02_AI-for-Science-Congress-2026_Sebastian-Eugen-Buzdugan.docx) | [PDF](pdf/02_AI-for-Science-Congress-2026_Codruta-Costea.pdf) · [DOCX](docx/02_AI-for-Science-Congress-2026_Codruta-Costea.docx) |
| MAIXPO 2026 | [PDF](pdf/03_MAIXPO-2026_Sebastian-Eugen-Buzdugan.pdf) · [DOCX](docx/03_MAIXPO-2026_Sebastian-Eugen-Buzdugan.docx) | [PDF](pdf/03_MAIXPO-2026_Codruta-Costea.pdf) · [DOCX](docx/03_MAIXPO-2026_Codruta-Costea.docx) |

- The combined file, [Combined_All-Six-Invitation-Requests.pdf](Combined_All-Six-Invitation-Requests.pdf), has 12 pages with one bookmark per document.
- [src/build_documents.py](src/build_documents.py) is the generator. It writes the DOCX files, then converts them to PDF with LibreOffice.

Every document uses the same layout: an A4 page with 20 mm side margins, a draft-status banner, an event header and the request letter on page 1, then an annex with verified facts, clickable source links, an outstanding-requirements checklist and a document-status note on page 2. Each footer reads "Not issued, signed or confirmed by the organizer".

## Research findings by event

### 1. AI Engineers Meetup Bangkok

- **Confirmed on the listing:** Saturday 26 Dec 2026, 10:00–13:00 ICT, at SCBX NEXT TECH, 4th Fl., Siam Paragon. The group is "AI Engineers in Thailand" and the host is shown as Ek C. It is a weekly free coffee meetup, and 26 December is the last date of the recurring series.
- **Not published:** a fee, capacity, attendee count, contact email, or any process for invitation letters or tickets. Meetup has no attendee-letter feature.
- **Sources:** the [event listing](https://www.meetup.com/th-th/coding-in-bangkok/events/nsntvtyjcqbjc/) and the [group page](https://www.meetup.com/coding-in-bangkok/).

### 2. AI for Science Congress 2026

- **Confirmed:** 27–30 Dec 2026 in Kuala Lumpur, Malaysia, co-located with ISPA, BDCloud, SocialCom, SustainCom and SpaCCS 2026. The "Sponsored and Organized by" section shows Zhengzhou University and Zhejiang Normal University. The contact address is `ai4science-2026@googlegroups.com`. The general chairs listed are Stephen Yau, M. Jamal Deen and Zhiyuan Chen.
- **Registration:** the [registration instructions](https://ieee-ai-for-science.org/2026/ispa/author/aure.php) describe only *Regular registration for paper authors*. The fee is USD 750 for IEEE members and USD 900 for non-members, the deadline is 30 Nov 2026, and payment is by bank transfer. **No non-author (attendee-only) registration is published.**
- **Not published:** a venue (the [venue page](https://ieee-ai-for-science.org/2026/venue.php) still shows the 2025 Zhengzhou content), the travel page (it returns an error), a programme, or any visa or invitation-letter procedure.
- **Branding:** the site's visible header logo file is named "ieee mb white" and its contact page says "IEEE AI for Science Congress 2026". The IEEE, IEEE CS and technical-committee sponsor logos and the "IEEE" page title are commented out in the page code. The draft therefore uses the event name the user gave and claims no IEEE sponsorship.
- **Sources:** the [website](https://ieee-ai-for-science.org/2026/) and the [contact page](https://ieee-ai-for-science.org/2026/contact.php).
- **What I could not find:** no template from this congress series, either for 2026 or from the earlier Cybermatics/ISPA editions. Other IEEE conferences publish visa-letter policies (for example [IEEE-USA guidelines](https://futureleaders.ieeeusa.org/register/visa-letter-requests/)), but those belong to other events and were **not** used as templates.

### 3. MAIXPO 2026

- **Confirmed:** "MAIXPO 2026 — AI Marketing Conference & Expo" is a hybrid event in Kuala Lumpur on 29–30 Dec 2026. The venue "will be announced once confirmed". Tickets are open: Standard early bird is USD 49 and VIP early bird is USD 99, both until 30 Oct 2026. Payment instructions are sent by email because online payment is "coming soon".
- **Invitation-letter procedure (FAQ):** letters are issued to **confirmed ticket holders**. After registering, you email `info@maixpo.com` with your full name, passport number, nationality and country of application. The draft follows this procedure exactly. Source: [maixpo.com/#faq](https://maixpo.com/#faq).
- **Discrepancies to raise before paying:**
  - The site says "Wednesday & Thursday", but 29–30 Dec 2026 fall on a **Tuesday and Wednesday**.
  - The FAQ calls it "the only October 2026 event in KL".
  - The site names **no legal entity or company address**.
  - The "Privacy" and "Terms" links open an email instead of a policy page.
  - All speakers are listed as TBA.

## Requirements still outstanding

| Event | Before sending the request | Before travel |
|---|---|---|
| Meetup | Each attendee RSVPs from their own Meetup account. Send the request through Meetup messaging. | Check that the 26 Dec session is still scheduled. Check Thailand entry rules for your nationality and complete the Thailand Digital Arrival Card. |
| AI for Science | Email the organizers to confirm that **attendee-only registration exists** for 29–30 Dec, and ask for its fee and an invitation-letter procedure. Before any bank transfer, confirm the bank details with the organizers by email. | Confirm the venue. Check Malaysia entry rules and complete the MDAC. |
| MAIXPO | **Buy and confirm tickets** (early bird ends 30 Oct 2026). Then email the request with passport details. Settle the weekday and entity questions above first. | Confirm the venue once it is announced. Check Malaysia entry rules and complete the MDAC. |

**For both attendees, in every document:** fill in address, email, phone, date, nationality, passport number and expiry, country of application and organization. Check that each name matches the passport exactly. The drafts use the names as you supplied them, without diacritics. Then sign above "(requester)", which is the attendee's own signature.

## Quality checks performed

- I rendered all 12 pages to images and inspected them. Each document fits on 2 pages with no text overflow, and margins and table widths are consistent.
- Every PDF is A4 (595 × 842 pt) with embedded fonts: Inter, Roboto, DM Sans and Bebas Neue, plus DejaVu Sans for the ☐ checkboxes.
- I checked each document's text for the attendee's own name (5 occurrences) and the correct companion name, and reviewed the full word list for spelling.
- The weekdays were computed: 26 Dec 2026 is a Saturday, 29 Dec is a Tuesday and 30 Dec is a Wednesday.
- Clickable links per PDF: 2 for the Meetup, 4 for the AI for Science Congress and 3 for MAIXPO.
- The DOCX files are the editable source. To get identical PDF output, install the same fonts; otherwise Word substitutes them.
