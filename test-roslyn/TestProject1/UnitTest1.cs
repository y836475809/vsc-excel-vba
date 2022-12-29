using ConsoleApp1;
using System;
using System.Collections.Generic;
using Xunit;

namespace TestProject1
{
    public class UnitTest1
    {
        [Fact]
        public async void Test1()
        {
            var class1Name = "test_class1.cls";
            var class1Code = Helper.getCode(class1Name);
            var mod1Name = "test_module1.bas";
            var mod1Code = Helper.getCode(mod1Name);
            var mc = new MyCodeAnalysis();
            mc.AddDocument(class1Name, class1Code);
            mc.AddDocument(mod1Name, mod1Code);
            var items = await mc.GetCompletions(mod1Name, mod1Code, Helper.getPosition(mod1Code, "p."));
            var exp = new List<CompletionItem>() {
                new CompletionItem(){
                    DisplayText = "Public Name As String",
                    CompletionText = "Name",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Mother As Person",
                    CompletionText = "Mother",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Sub SayHello(val1 As Object, val2 As Object)",
                    CompletionText = "SayHello",
                    Description = "<member name=\"M:Person.SayHello(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    ReturnType = "Void",
                    Kind = "Method",
                },
            };
            Assert.Equal(exp, items);
        }

        [Fact]
        public async void Test2()
        {
            var class1Name = "test_class1.cls";
            var class1Code = Helper.getCode(class1Name);
            var mod1Name = "test_module1.bas";
            var mod1Code = Helper.getCode(mod1Name);
            var mc = new MyCodeAnalysis();
            mc.AddDocument(class1Name, class1Code);
            mc.AddDocument(mod1Name, mod1Code);

            var items1 = await mc.GetCompletions(mod1Name, mod1Code, Helper.getPosition(mod1Code, "p."));
            var exp = new List<CompletionItem>() {
                new CompletionItem(){
                    DisplayText = "Public Name As String",
                    CompletionText = "Name",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Mother As Person",
                    CompletionText = "Mother",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Sub SayHello(val1 As Object, val2 As Object)",
                    CompletionText = "SayHello",
                    Description = "<member name=\"M:Person.SayHello(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    ReturnType = "Void",
                    Kind = "Method",
                },
            };
            Assert.Equal(exp, items1);


            var class2Code = Helper.getCode("test_class2.cls");
            mc.ChangeDocument(class1Name, class2Code);
            var items2 = await mc.GetCompletions(mod1Name, mod1Code, Helper.getPosition(mod1Code, "p."));
            var exp2 = new List<CompletionItem>() {
                new CompletionItem(){
                    DisplayText = "Public Name As String",
                    CompletionText = "Name",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Mother As Person",
                    CompletionText = "Mother",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Sub SayHello2(val1 As Object, val2 As Object)",
                    CompletionText = "SayHello2",
                    Description = "<member name=\"M:Person.SayHello2(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    ReturnType = "Void",
                    Kind = "Method",
                },
            };
            Assert.Equal(exp2, items2);
        }

        [Fact]
        public async void Test3()
        {
            var class1Name = "test_class1.cls";
            var class1Code = Helper.getCode(class1Name);
            var mod1Name = "test_module1.bas";
            var mod1Code = Helper.getCode(mod1Name);
            var mc = new MyCodeAnalysis();
            mc.AddDocument(class1Name, class1Code);
            mc.AddDocument(mod1Name, mod1Code);
            var items = await mc.GetCompletions(mod1Name, mod1Code, Helper.getPosition(mod1Code, "p."));
            var exp = new List<CompletionItem>() {
                new CompletionItem(){
                    DisplayText = "Public Name As String",
                    CompletionText = "Name",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Mother As Person",
                    CompletionText = "Mother",
                    Description = "",
                    ReturnType = "",
                    Kind = "Field",
                },
                new CompletionItem(){
                    DisplayText = "Public Sub SayHello(val1 As Object, val2 As Object)",
                    CompletionText = "SayHello",
                    Description = "<member name=\"M:Person.SayHello(System.Object,System.Object)\">\r\n <summary>\r\n  テストメッセージ\r\n </summary>\r\n <param name='val1'></param>\r\n <param name='val2'></param>\r\n <returns></returns>\r\n</member>",
                    ReturnType = "Void",
                    Kind = "Method",
                },
            };
            Assert.Equal(exp, items);

            mc.DeleteDocument(class1Name);
            var items2 = await mc.GetCompletions(mod1Name, mod1Code, Helper.getPosition(mod1Code, "p."));
            Assert.Equal(new List<CompletionItem>() { }, items2);
        }
    }
}
