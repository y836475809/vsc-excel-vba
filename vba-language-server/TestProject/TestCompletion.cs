using VBACodeAnalysis;
using System;
using System.Collections.Generic;
using Xunit;
using System.Threading.Tasks;

namespace TestProject {
    public class TestCompletion
    {
        [Fact]
        public async Task Test1()
        {
            var class1Name = "test_class1.cls";
            var class1Code = Helper.getCode(class1Name);
            var mod1Name = "test_module1.bas";
            var mod1Code = Helper.getCode(mod1Name);
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.AddDocument(class1Name, class1Code);
            vbaca.AddDocument(mod1Name, mod1Code);
            var items = await vbaca.GetCompletions(mod1Name, 11, 6);
            var exp = new List<VBACompletionItem>() {
                new(){
                    Label = "Name",
                    Display = "Public Name As String",
					Doc = "",
                    Kind = "Field",
				},
                new(){
					Label = "Mother",
					Display = "Public Mother As Person",
					Doc = "",
                    Kind = "Field",
				},
                new(){
					Label = "SayHello",
					Display = "Public Sub SayHello(val1 As Object, val2 As Object)",
					Doc = "<member name=\"M:Person.SayHello(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    Kind = "Method",
                }
            };
			Helper.AssertCompletionItem(exp, items);
        }

        [Fact]
        public async Task Test2()
        {
            var class1Name = "test_class1.cls";
            var class1Code = Helper.getCode(class1Name);
            var mod1Name = "test_module1.bas";
            var mod1Code = Helper.getCode(mod1Name);
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.AddDocument(class1Name, class1Code);
            vbaca.AddDocument(mod1Name, mod1Code);

            var items1 = await vbaca.GetCompletions(mod1Name, 11,6);
            var exp = new List<VBACompletionItem>() {
                new(){
					Label = "Name",
					Display = "Public Name As String",
					Doc = "",
                    Kind = "Field",
                },
                new(){
					Label = "Mother",
					Display = "Public Mother As Person",
					Doc = "",
                    Kind = "Field",
                },
                new(){
					Label = "SayHello",
					Display = "Public Sub SayHello(val1 As Object, val2 As Object)",
					Doc = "<member name=\"M:Person.SayHello(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    Kind = "Method",
                }
            };
            Helper.AssertCompletionItem(exp, items1);


            var class2Code = Helper.getCode("test_class2.cls");
            vbaca.ChangeDocument(class1Name, class2Code);
            var items2 = await vbaca.GetCompletions(mod1Name, 11,6);
            var exp2 = new List<VBACompletionItem>() {
                new(){
					Label = "Name",
					Display = "Public Name As String",
					Doc = "",
                    Kind =  "Field",
                },
                new(){
					Label = "Mother",
					Display = "Public Mother As Person",
					Doc = "",
                    Kind = "Field",
				},
                new(){
					Label = "SayHello2",
					Display = "Public Sub SayHello2(val1 As Object, val2 As Object)",
					Doc = "<member name=\"M:Person.SayHello2(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    Kind = "Method",
                }
            };
            Helper.AssertCompletionItem(exp2, items2);
        }

        [Fact]
        public async Task Test3()
        {
            var class1Name = "test_class1.cls";
            var class1Code = Helper.getCode(class1Name);
            var mod1Name = "test_module1.bas";
            var mod1Code = Helper.getCode(mod1Name);
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.AddDocument(class1Name, class1Code);
            vbaca.AddDocument(mod1Name, mod1Code);
            var items = await vbaca.GetCompletions(mod1Name, 11,6);
            var exp = new List<VBACompletionItem>() {
                new(){
					Label = "Name",
					Display = "Public Name As String",
					Doc = "",
                    Kind = "Field",
                },
                new(){
					Label = "Mother",
					Display = "Public Mother As Person",
					Doc = "",
                    Kind = "Field",
                },
                new(){
					Label = "SayHello",
					Display = "Public Sub SayHello(val1 As Object, val2 As Object)",
					Doc = "<member name=\"M:Person.SayHello(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    Kind = "Method",
                }
            };
            Helper.AssertCompletionItem(exp, items);

            vbaca.DeleteDocument(class1Name);
            var items2 = await vbaca.GetCompletions(mod1Name, 11,6);
            Assert.Empty(items2);
        }
    }
}
