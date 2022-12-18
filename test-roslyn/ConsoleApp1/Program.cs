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

var code = @"using System;
public class MyClass
{
    public static void MyMethod(int value)
    {
        var aa_b = new TestClass2();
        aa_b.
    }
}";
            var projectInfo = ProjectInfo.Create(ProjectId.CreateNewId(), VersionStamp.Create(), "MyProject", "MyProject", LanguageNames.CSharp).
               WithMetadataReferences(new[]
               { 
                   MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
               });

            var project = workspace.AddProject(projectInfo);
            //SourceText sourcetext = SourceText.From(code);
            workspace.AddDocument(project.Id, "TestC.cs", SourceText.From(code_class));
            var document = workspace.AddDocument(project.Id, "MyFile.cs", SourceText.From(code));
            var position = code.LastIndexOf("aa_b.") + 5;
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
