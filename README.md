# Description

This plugin allows content in a Google docs document to be exported to [Anki](https://apps.ankiweb.net) flashcards.

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

# Usage

## Format your content

### Front of cards
All headings of the lowest level in a heading "tree" will be used as the 
content for the Front of the card.

### Back of cards
All content under the lowest leveled headings will be used as content.

### Example
This content:<br>
<img width="705" alt="image" src="https://user-images.githubusercontent.com/1998726/213292753-a6e6440c-bb0f-4dae-b508-f279b0c4d250.png">

Will generate these Notes:<br>
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



