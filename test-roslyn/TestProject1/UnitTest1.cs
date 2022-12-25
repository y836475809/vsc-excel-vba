using System;
using Xunit;

namespace TestProject1 {
    public class UnitTest1 {
        [Fact]
        public void Test1() {
            var mc = new ConsoleApp1.MyCodeAnalysis();
            Assert.Equal(10, mc.p1());
        }
    }
}
