# Description

This plugin allows content in a Google docs document to be exported to [Anki](https://apps.ankiweb.net) flashcards.

![image](https://user-images.githubusercontent.com/1998726/213298694-54ddb150-3e4b-4bc6-be22-95b78f4ab912.png)


# Features

- Scans the selected part of the document, or the whole document if nothing is selected.

## Supported file formats:
  - .apkg - for import directly into Anki
  - zipped .csv files for further processing
## Supported content:
 - Text
 - Images (will not display in the sidebar preview)
 - Lists
    - Nested lists work
    - Ordered (with numbers) does not

# Installation

## Host bundle.js
- You need to host the **bundle.js** file on a server somewhere.
  - Google allows free hosting of up to 5GB of traffic a month (which should be plenty for this use case)
  - See [Google's documentation](https://cloud.google.com/storage/docs/hosting-static-website)
- Enter the URL of the hosted bundle.js into **sidebar.html** (line 211)
  - "<script src="[URL TO REACH bundle.js]"></script>"

## Install code into document
- Create or open a Google docs document.
- Open [Extensions]/[Apps scripts]
- Paste code from **code.gs** into the code file that is automatically created (replace the existing code).
- Create a new HTML file.
  - Name it "Sidebar" (html extension will be added automatically).
  - Paste code from **Sidebar.html** into the file you created.
- Close the document and re-open it.
- Open Sidebar from [Extensions]/[Anki-Export]/[Show in Sidebar].
  - Give the script the permissions it asks for.

# Usage

## Format your content

### Front of cards
- All headings of the lowest level in a heading "tree" will be used as the 
content for the Front of the card.
- Headings above the lowest levels will be ignored but can still be useful to structure the content in the document.

**Example:**
- Heading 1 (ignored)
  - Heading 2 (ignored)
    - Heading 3 (Front)
      - Content (Back) 

See detailed example furher down.

### Back of cards
All content under the lowest leveled headings will be used as content.

### Example
**This content:**<br>
<img width="705" alt="image" src="https://user-images.githubusercontent.com/1998726/213292753-a6e6440c-bb0f-4dae-b508-f279b0c4d250.png">

**Will generate these Notes:**<br>
<img width="317" alt="image" src="https://user-images.githubusercontent.com/1998726/213294441-8ec304d6-5fe3-4b91-8d73-a4931b945665.png">
 
## Find Notes
- Open Sidebar from [Extensions]/[Anki-Export]/[Show in Sidebar].
- Press "Find cards".
<img width="345" alt="image" src="https://user-images.githubusercontent.com/1998726/213293930-63417f48-81ed-4783-8d01-9ef153f743b1.png">

- A progress bar will be displayed while the content is processed.
<img width="330" alt="image" src="https://user-images.githubusercontent.com/1998726/213294026-075c5613-7b4b-47b0-95a9-966f6c37da5f.png">

- The cards will be listed under "Front".
- Select a card to preview the back/content of it.
<img width="322" alt="image" src="https://user-images.githubusercontent.com/1998726/213294805-ce22ce2f-d314-4ec0-8c07-89d1ce21c389.png">

## Export
- Press one of the two export buttons at the bottom of the sidebar.
<img width="322" alt="image" src="https://user-images.githubusercontent.com/1998726/213296306-c11b7553-dfd5-4d43-adff-4810e6d02544.png">

# Detailed content example
The content is formatted as follows:<br>
[Heading level 1]<br>
Surgery<br>
[Heading level 2]<br>
Cause of abdominal distension in inflammatory conditions of the abdomen?
- Intestinal paralysis causes swollen intestines
  - Gas and fluid in the intestines
<br>

[Heading level 1]<br>
Ileus<br>
[Heading level 2]<br>
What causes pain in case of inflammation of the abdomen - Peritoneum Viscerale or parietale?
- peritoneum parietale
<br>
[Heading level 2]<br>
What drives the pain response in ileus?

- It is the bowel that drives the pain
- Palpation will not make much difference 
  - Patient will experience discomfort on palpation
<br>
Pain will be non-specific in localisation



