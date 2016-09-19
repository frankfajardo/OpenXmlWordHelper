# OpenXmlWordHelper
Just a static class with helper methods for updating Word using OpenXml

### Sample Usage

- Replacing all merge fields named *PrintDate* 
  ```
  using (WordprocessingDocument document = WordprocessingDocument.Open(docPath, true))
  {
     document.GetMergeFields("PrintDate").ReplaceWithText(DateTime.Now.ToString("dd MMMM yyyy"));
     document.MainDocumentPart.Document.Save();
  }
  ```

- Replacing merge fields contained in a `Paragraph`
  ```
  thatParagraph.GetMergeFields("FirstName").ReplaceWithText(firstName);
  thatParagraph.GetMergeFields("LastName").ReplaceWithText(lastName);
  ```


- Replacing multiple merge fields with the same name, eg. a letter recipient's name and address lines, which can have variable number of lines depending on the data available
  ```
  List<string> nameAndAddressList = GetRecipientNameAndAddressLines();
  using (WordprocessingDocument document = WordprocessingDocument.Open(docPath, true))
  {
     document.GetMergeFields("RecipientNameAndAddress").ReplaceWithText(nameAndAddressList, removeExcess: true);
     document.MainDocumentPart.Document.Save();
  }
  ```


