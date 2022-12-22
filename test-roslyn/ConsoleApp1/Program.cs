using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Completion;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Recommendations;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1 {
    class Program {
        [Obsolete]
        static async Task Main(string[] args) {
            Console.WriteLine("Hello World!");

            var host = MefHostServices.Create(MefHostServices.DefaultAssemblies);
            var workspace = new AdhocWorkspace(host);
            var code_class = @"using System;
public class TestClass
{
     public int TetsMethod(int value){
        return value;
    }
     public string TetsMethodStr(string value){
        return value;
    }
}
public class TestClass2
{
  /// <summary>
  /// サンプルコード21
  /// </summary>
  /// <param name='hoge'>第１引数</param>
  /// <param name='fuga'>第２引数</param>
  /// <returns>関数の結果</returns>
     public int TetsMethod(int value){
        return value;
    }
  /// <summary>
  /// サンプルコード22
  /// </summary>
  /// <param name='hoge'>第１引数</param>
  /// <param name='fuga'>第２引数</param>
  /// <returns>関数の結果</returns>
     public string TetsMethodStr(string value){
        return value;
    }
}
";

var code_class_mod = @"using System;
public class TestClass
{
     public int TetsMethod(int value){
        return value;
    }
     public string TetsMethodStr(string value){
        return value;
    }
}
public class TestClass2
{
  /// <summary>
  /// サンプルコード21
  /// </summary>
  /// <param name='hoge'>第１引数</param>
  /// <param name='fuga'>第２引数</param>
  /// <returns>関数の結果</returns>
     public int TetsMethodMod(int value){
        return value;
    }
  /// <summary>
  /// サンプルコード22
  /// </summary>
  /// <param name='hoge'>第１引数</param>
  /// <param name='fuga'>第２引数</param>
  /// <returns>関数の結果</returns>
     public string TetsMethodStrMod(string value){
        return value;
    }
}
";

            var code = @"using System;
public class MyClass
{
    public static void MyMethod(int value)
    {
        var aa_b = new TestClass2();
        aa_b.Tets
    }
}";
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.CSharp).
               WithMetadataReferences(new[]
               { 
                   MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
               });

            var project = workspace.AddProject(projectInfo);
            //SourceText sourcetext = SourceText.From(code);

            var doc_class = workspace.AddDocument(project.Id, "TestC.cs", SourceText.From(code_class));
            var document = workspace.AddDocument(project.Id, "MyFile.cs", SourceText.From(code));
            var position = code.LastIndexOf("aa_b.Tets") + 9;
            var completionService = CompletionService.GetService(document);
            //Microsoft.CodeAnalysis.Options.DocumentOptionSet dopset;
            //dopset.
            var results = await completionService.GetCompletionsAsync(document, position);
            var mmp = results.SuggestionModeItem;
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
            Console.WriteLine("-------------------");
            {
               
                var solution = workspace.CurrentSolution;
                //var documentIds = solution.GetDocument(doc_class.Id);
                //var documents = project.Documents;
                var doc = workspace.CurrentSolution.GetDocument(doc_class.Id);
                //var sourceText = await doc.GetTextAsync();
                //solution = doc.WithText(SourceText.From(code_class_mod)).Project.Solution;
                solution = solution.WithDocumentText(doc_class.Id, SourceText.From(code_class_mod));
                //var e = workspace.TryApplyChanges(solution);

                ////SourceText sourceText = SourceText.From(code_class_mod);
                ////Solution newSolution = doc_class.Project.Solution.WithDocumentText(doc_class.Id, sourceText);
                //var doc = workspace.CurrentSolution.GetDocument(doc_class.Id);
                //var np = doc.WithText(SourceText.From(code_class_mod)).Project;
                //np = np.GetDocument(document.Id).WithText(SourceText.From(code)).Project;
                ////Document doc = project.GetDocument(doc_class.Id);
                ////var ndoc = doc_class;
                ////project = document.Project;
                //var e = workspace.TryApplyChanges(np.Solution);
                //var e = workspace.TryApplyChanges(
                //    doc.WithText(SourceText.From(code_class_mod)).Project.Solution);
                //var e2 = workspace.TryApplyChanges(document.Project.Solution);
                //var e =  workspace.TryApplyChanges(doc_class.WithText(SourceText.From(code_class_mod)).Project.Solution);
                //var e2 = workspace.TryApplyChanges(document.Project.Solution);
                solution = solution.GetDocument(document.Id).
                    WithText(SourceText.From(code)).Project.Solution;
                //var newDoc = doc.WithText(SourceText.From(code));
                //solution = newDoc.Project.Solution;
                workspace.TryApplyChanges(solution);
                //var currentDoc = _workspace.CurrentSolution.GetDocument(docId);
                document = workspace.CurrentSolution.GetDocument(document.Id);

                //solutionWorkspace.TryApplyChanges(
                //    solutionDocument.WithText(SourceText.From(editorText)).Project.Solution);
                //var newDoc = document.WithText(SourceText.From(code_class_mod));
                //workspace.TryApplyChanges(newDoc.Project.Solution);
                //var document2 = workspace.CurrentSolution.GetDocument(document.Id);
                //workspace.AddDocument(project.Id, "TestC.cs", SourceText.From(code_class_mod));
                //var document2 = workspace.AddDocument(project.Id, "MyFile.cs", SourceText.From(code));
                //workspace.TryApplyChanges(document.Project.Solution);
                //var document2 = workspace.AddDocument(project.Id, "MyFile.cs", SourceText.From(code));
                var position2 = code.LastIndexOf("aa_b.Tets") + 9;
                var completionService2 = CompletionService.GetService(document);
                //Microsoft.CodeAnalysis.Options.DocumentOptionSet dopset;
                //dopset.
                var results2 = await completionService2.GetCompletionsAsync(document, position2);
                foreach (var i in results2.ItemsList) {
                    var dd = i.InlineDescription;
                    Console.WriteLine(i.DisplayText);

                    foreach (var prop in i.Properties) {
                        Console.Write($"{prop.Key}:{prop.Value} ");
                    }

                    Console.WriteLine();
                    foreach (var tag in i.Tags) {
                        Console.Write($"{tag} ");
                    }

                    Console.WriteLine();
                    Console.WriteLine();
                }

                {
                    var model = document.GetSemanticModelAsync().Result;
                    var symbols = Recommender.GetRecommendedSymbolsAtPosition(model, position2, workspace);
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

            //var syntaxTree = CSharpSyntaxTree.ParseText(code_class);
            //var mscorlib = MetadataReference.CreateFromFile(typeof(object).Assembly.Location);
            //var compilation = CSharpCompilation.Create("MyCompilation", new[] { syntaxTree }, new[] { mscorlib });
            //var semanticModel = compilation.GetSemanticModel(syntaxTree);
            ////semanticModel.GetDeclaredSymbol()

            //var syntaxRoot = syntaxTree.GetRoot();
            //var classNode = syntaxRoot.DescendantNodes().OfType<MethodDeclarationSyntax>();
            //var symbolInfo = semanticModel.GetDeclaredSymbol(classNode.First());
            //var parts = symbolInfo.ToDisplayParts();
            //var disp_strs = symbolInfo.ToDisplayString();
            ////var containingType = symbol.ContainingType;
            ////var overloads = containingType.TypeArguments;

            //var cp = await project.GetCompilationAsync();
            ////project.Documents
            ////project.GetSyntaxRootAsync
            //var mm = syntaxRoot.DescendantNodes().OfType<MethodDeclarationSyntax>();
            //var pp = mm.Where(x => x.Modifiers.ToString().Contains("public"));
            ////var methods = from m in syntaxRoot.DescendantNodes().OfType<MethodDeclarationSyntax>() where m.Modifiers.ToString().Contains("public") select m;

            ////var symbols = await SymbolFinder.FindSourceDeclarationsAsync(project, x => x.Equals(typeName));
            ////return symbols.Where(x => x.Kind == SymbolKind.NamedType).Cast<INamedTypeSymbol>().ToList();
            ////var members = symbol.GetMembers().Where(x => x.Name.Equals("MyMethod"));
            ////classNode.First().
            //var methodNode = (MethodDeclarationSyntax)classNode.First();
            ////methodNode.Identifier.
            //string modelClassName = string.Empty;

            //foreach (var param in methodNode.ParameterList.Parameters) {
            //    var metaDataName = document.GetSemanticModelAsync().Result.GetDeclaredSymbol(param).ToDisplayString();
            //    //'document' is the current 'Microsoft.CodeAnalysis.Document' object
            //    var members = document.Project.GetCompilationAsync().Result.GetTypeByMetadataName(metaDataName).GetMembers();
            //    var props = (members.OfType<IPropertySymbol>());

            //    //now 'props' contains the list of properties from my type, 'Type1'
            //    foreach (var prop in props) {
            //        //some logic to do something on each proerty
            //    }
            //}
            //var m = 0;
            ////var classModel = (ITypeSymbol)semanticModel.GetDeclaredSymbol(classNode);
            ////var firstAttribute = classModel.GetAttributes().First();
        }
    }
}
