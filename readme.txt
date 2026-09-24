===========================================================================
                               Link Checker
Author: Thomas Tang
Version: 1.0
===========================================================================

****** Version 1.0 (2026-02-18) ******
(A) Background
	To minimise the time and effort in checking active links to the following regulatory disclosures to platforms, this code automates the manual process of clicking each link and checking PDF content.
	- EU PRIIPS KID
	- UK PRIIPS KID
	- UK MiFID KIID


(B) Prerequisites
	- All lists must be in .xlsx format.
	- All lists must have the same headers in following order (see first tab of the example EU KID Links file as example):
		1. Doc Type
		2. ISIN
		3. Name (Preferably it's a hard-coded value, not lookup results from formulas. Otherwise, name may not be populated in output summary, though this would not impact the overall checks.)
		4. Language
		5. New URL


(C) Execution
	1. Move all Excel lists to the module directory first.
	2. Run main.py on PyCharm.
	3. You would be prompted to input the name of your Excel list. Follow instructions on screen.
	4. When the checking is completed, the checking results would be shown in a new tab Output in your list file.
	5. Consider removing the downloaded PDFs after use to save storage memory (Optional) 

	N.B. Returned PDFs would be downloaded in the process and saved at a designated folder (can be seen from __init__.py).


(D) Interpreting Check Results
	status_code, valid_link
	- A status code of 200 indicates the link concerned works. Anything other than that means the link doesn’t work and probably needs renewal.

	doc_type
	- What the link returns, provided that the link is valid. The link may need review if it returns something other than a pdf.

	last_upload_date
	- When the document was last uploaded. This does not indicate the version of the uploaded document. But if the document was last uploaded a long time ago, there is a high chance that the document is not the latest.

	as_of_date, pdf_isin
	- As of date and ISIN extracted from the returned PDF.