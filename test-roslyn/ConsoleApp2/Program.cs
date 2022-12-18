using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Recommendations;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Threading.Tasks;

namespace ConsoleApp2 {
    class Program {
        [Obsolete]
        static async Task Main(string[] args) {
            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            var workspace = new AdhocWorkspace(host);
            var code_module1 = @"
''' <summary>
'''  テストメッセージ
''' </summary>
''' <param name='val1'></param>
''' <param name='val2'></param>
''' <returns></returns>
Sub Sample1()
    Range('A1') = 'tanaka'
End Sub

Sub call1()
    Dim p As Person
    Set p = New Person
    p.
End Sub
";

            var code_class1 = @"
Option Explicit
Public Class Person
' メンバ変数
Public Name As String
Private Age As Long
Public Mother As Person

' メソッド
Public Sub SayHello()
    MsgBox ""Hello!""
End Sub
End Class
";
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), 
                "MyProject", "MyProject", LanguageNames.VisualBasic).
               WithMetadataReferences(new[]
               {
                   MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
               });

            var project = workspace.AddProject(projectInfo);
            //SourceText sourcetext = SourceText.From(code);
            workspace.AddDocument(project.Id, "Person.cls", SourceText.From(code_class1));
            var document = workspace.AddDocument(project.Id, "MyFile.bas", SourceText.From(code_module1));
            var position = code_module1.LastIndexOf("p.") + 2;
            var completionService = CompletionService.GetService(document);
            //Microsoft.CodeAnalysis.Options.DocumentOptionSet dopset;
            //dopset.
            var results = await completionService.GetCompletionsAsync(document, position);
            foreach (var i in results.ItemsList) {
                var dd = i.InlineDescription;
                Console.WriteLine(i.DisplayText);

                foreach (var prop in i.Properties) {
                    Console.Write($"{prop.Key}:{prop.Value} ");
                    //document.GetSemanticModelAsync().Result.SyntaxTree.GetRoot().
                }

                Console.WriteLine();
                foreach (var tag in i.Tags) {
                    Console.Write($"{tag} ");
                }

                Console.WriteLine();
                Console.WriteLine();
            }

            {
                //var completionService2 = CompletionService.GetService(document);
                //var charCompletion = GetCompletionTrigger('.');
                //var data = await completionService.GetCompletionsAsync(document, position).ConfigureAwait(false);
                //if (data == null || data.Items.Any() == false)
                //    return new List<ICompletionData>();
                var model = document.GetSemanticModelAsync().Result;
                var symbols = Recommender.GetRecommendedSymbolsAtPosition(model, position, workspace);
                foreach (var item in symbols) {
                    var n = item.Name;
                    var k = item.Kind;
                    var md = item.MetadataName;
                    Console.WriteLine(item.ToDisplayString());
                    Console.WriteLine(item.GetDocumentationCommentXml());
                    Console.WriteLine();
                }
                var ppp = 0;
            }
        }
    }
}
