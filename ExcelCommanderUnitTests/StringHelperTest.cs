﻿using ExcelCommander.Base;

namespace ExcelCommanderUnitTests
{
    public class StringHelperTest
    {
        [Fact]
        public void ParametersSplitShouldConsiderQuotes()
        {
            Assert.Equal(2, "hello world".SplitParameters().Length);
            Assert.Equal(2, "Command \"Argument 1\"".SplitParameters().Length);
            Assert.Equal("Argument 2", "Command \"Argument 1\" \"Argument 2\"".SplitParameters().Last());
        }
        [Fact]
        public void ParametersSplitShouldConsiderIncludeQuotes()
        {
            Assert.Equal(2, "hello world".SplitParameters(true).Length);
            Assert.Equal(2, "Command \"Argument 1\"".SplitParameters(true).Length);
            Assert.Equal("\"Argument 2\"", "Command \"Argument 1\" \"Argument 2\"".SplitParameters(true).Last());
        }
    }
}
