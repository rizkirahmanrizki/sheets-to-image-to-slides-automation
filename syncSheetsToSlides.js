/**
 * Sheets → Slides Automation
 * ------------------------------------------------------------
 * This script exports specific ranges from a spreadsheet,
 * converts them into images, and inserts them into slides.
 *
 * Use case:
 * - automated reporting
 * - dashboard snapshot updates
 * - scheduled slide refresh
 *
 * NOTE:
 * All IDs are placeholders. Replace with your own.
 */

function syncSheetsToSlides() {

  const CONFIG = {
    SPREADSHEET_ID: "SPREADSHEET_ID_PLACEHOLDER",
    PRESENTATION_ID: "PRESENTATION_ID_PLACEHOLDER"
  };

  const token = ScriptApp.getOAuthToken();

  /**
   * Jobs configuration:
   * Each job defines:
   * - sheet (gid)
   * - range
   * - slide index
   * - label
   */
  const jobs = [
    { gid: "SHEET_GID_1", range: "A1:D20", slideIndex: 0, name: "Section 1" },
    { gid: "SHEET_GID_2", range: "A1:D20", slideIndex: 1, name: "Section 2" },
    { gid: "SHEET_GID_3", range: "A1:D20", slideIndex: 2, name: "Section 3" }
  ];

  const presentation = SlidesApp.openById(CONFIG.PRESENTATION_ID);
  const slideWidth = presentation.getPageWidth();
  const slideHeight = presentation.getPageHeight();

  jobs.forEach(job => {

    try {

      /* -------------------------------------------------------
         STEP 1: Export sheet range as PDF
      ------------------------------------------------------- */
      const exportUrl =
        `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/export?` +
        `exportFormat=pdf&gid=${job.gid}&range=${job.range}` +
        `&size=A4&portrait=false&fitw=true&gridlines=false`;

      const pdfBlob = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: "Bearer " + token }
      }).getBlob().setName(`${job.name}.pdf`);

      /* -------------------------------------------------------
         STEP 2: Convert PDF → Image (via thumbnail)
      ------------------------------------------------------- */
      const tempFile = DriveApp.createFile(pdfBlob);

      Utilities.sleep(2000); // ensure thumbnail is ready

      const thumbnailUrl =
        `https://drive.google.com/thumbnail?sz=w1000&id=${tempFile.getId()}`;

      const imageBlob = UrlFetchApp.fetch(thumbnailUrl, {
        headers: { Authorization: "Bearer " + token }
      }).getBlob().setName(`${job.name}.png`);

      const slide = presentation.getSlides()[job.slideIndex];

      /* -------------------------------------------------------
         STEP 3: Clear existing images
      ------------------------------------------------------- */
      slide.getPageElements().forEach(el => {
        if (el.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
          el.remove();
        }
      });

      /* -------------------------------------------------------
         STEP 4: Insert and scale image
      ------------------------------------------------------- */
      const image = slide.insertImage(imageBlob);

      const maxWidth = slideWidth - 40;
      const maxHeight = slideHeight - 40;

      const scale = Math.min(
        maxWidth / image.getWidth(),
        maxHeight / image.getHeight()
      );

      image
        .setWidth(image.getWidth() * scale)
        .setHeight(image.getHeight() * scale)
        .setLeft((slideWidth - image.getWidth()) / 2)
        .setTop((slideHeight - image.getHeight()) / 2)
        .sendToBack();

      /* -------------------------------------------------------
         STEP 5: Cleanup temp file
      ------------------------------------------------------- */
      tempFile.setTrashed(true);

      Logger.log(`Inserted: ${job.name}`);

    } catch (err) {
      Logger.log(`Error processing ${job.name}: ${err}`);
    }

  });

  Logger.log("All slides updated successfully.");

}
