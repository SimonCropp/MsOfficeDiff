public class SourceDocumentsVisibleTests
{
    [Test]
    public async Task SourceDocumentsRemainOpenAfterCompare()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Skip.Test("Microsoft Word is not installed");
        }

        dynamic word = Activator.CreateInstance(wordType!)!;
        try
        {
            word.DisplayAlerts = 0;
            word.Options.SaveInterval = 0;

            string tmp1, tmp2;
            var doc1 = Word.Open(word, ProjectFiles.input_temp_docx.FullPath, out tmp1);
            var doc2 = Word.Open(word, ProjectFiles.input_target_docx.FullPath, out tmp2);

            // Verify original files are not locked while open in Word
            using (var fs = new FileStream(ProjectFiles.input_temp_docx.FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
            }
            using (var fs = new FileStream(ProjectFiles.input_target_docx.FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
            }

            var compare = Word.LaunchCompare(word, doc1, doc2, tmp1, tmp2);

            word.Visible = true;

            // Non-quiet mode: ShowSourceDocuments should be set to both (3)
            Word.ApplyQuiet(false, word);
            await Assert.That((int) word.ActiveWindow.ShowSourceDocuments).IsEqualTo(3);

            // Quiet mode: ShowSourceDocuments should be set to none (0)
            Word.ApplyQuiet(true, word);
            await Assert.That((int) word.ActiveWindow.ShowSourceDocuments).IsEqualTo(0);

            compare.Saved = true;
            word.Quit(SaveChanges: false);
        }
        finally
        {
            Marshal.ReleaseComObject(word);
        }
    }
}
