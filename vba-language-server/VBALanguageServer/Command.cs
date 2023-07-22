using VBACodeAnalysis;
using System;
using System.Collections.Generic;

namespace VBALanguageServer {
    class Command {
        public string id { get; set; }

        public List<string> filepaths { get; set; }

        public int line { get; set; }
        public int chara { get; set; }

        public string text { get; set; }
    }

    public class DocumentAddedEventArgs : EventArgs {
        public List<string> FilePaths { get; set; }

        public DocumentAddedEventArgs(List<string> FilePaths) {
            this.FilePaths = FilePaths;
        }
    }

    public class DocumentDeletedEventArgs : EventArgs {
        public List<string> FilePaths { get; set; }

        public DocumentDeletedEventArgs(List<string> FilePaths) {
            this.FilePaths = FilePaths;
        }
    }

    public class DocumentRenamedEventArgs : EventArgs {
        public string OldFilePath { get; set; }
        public string NewFilePath { get; set; }

        public DocumentRenamedEventArgs(string OldFilePath, string NewFilePath) {
            this.OldFilePath = OldFilePath;
            this.NewFilePath = NewFilePath;
        }
    }

    public class DocumentChangedEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }

        public DocumentChangedEventArgs(string FilePath, string Text) {
            this.FilePath = FilePath;
            this.Text = Text;
        }
    }

    public class CompletionEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }
        public int Line { get; set; }
        public int Chara { get; set; }
        public List<CompletionItem> Items { get; set; }

        public CompletionEventArgs(string FilePath, string Text, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Line = Line;
            this.Chara = Chara;
        }
    }

    public class DefinitionEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }
        public int Line { get; set; }
        public int Chara { get; set; }

        public List<DefinitionItem> Items { get; set; }

        public DefinitionEventArgs(string FilePath, string Text, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Line = Line;
            this.Chara = Chara;
        }
    }

    public class DiagnosticEventArgs : EventArgs {
        public string FilePath { get; set; }

        public List<DiagnosticItem> Items { get; set; }

        public DiagnosticEventArgs(string FilePath) {
            this.FilePath = FilePath;
        }
    }

    public class DebugEventArgs : EventArgs {
        public string Text { get; set; }
        public DebugEventArgs() {
        }
    }

    public class ReferencesEventArgs : EventArgs {
        public string FilePath { get; set; }
        public int Line { get; set; }
        public int Chara { get; set; }
        public List<ReferenceItem> Items { get; set; }

        public ReferencesEventArgs(string FilePath, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Line = Line;
            this.Chara = Chara;
        }
    }

    public class SignatureHelpEventArgs : EventArgs {
        public string FilePath { get; set; }
        public string Text { get; set; }
        public int Line { get; set; }
        public int Chara { get; set; }
        public List<SignatureHelpItem> Items { get; set; }

        public SignatureHelpEventArgs(string FilePath, string Text, int Line, int Chara) {
            this.FilePath = FilePath;
            this.Text = Text;
            this.Line = Line;
            this.Chara = Chara;
        }
    }
}
