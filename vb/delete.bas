Attribute VB_Name = "delete"
Sub delete_page()
    Set doc = ActiveDocument
    
    'Identifying the current page
    Set currentPage = doc.Bookmarks.Item("\Page")
    
    'getting the range to be deleted
    Set currentRange = currentPage.Range
    
    'deleting the current page
    currentRange.delete
    
End Sub
