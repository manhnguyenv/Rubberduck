using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class VariableTypeNotDeclaredInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_MultipleParams()
        {
            const string inputCode =
                @"Sub Foo(arg1, arg2)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_DoesNotReturnResult_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_SomeTypesNotDeclared_Parameters()
        {
            const string inputCode =
                @"Sub Foo(arg1, arg2 As String)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_SomeTypesNotDeclared_Variables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
    Dim var2 As Date
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
    Dim var2
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_DoesNotReturnResult_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As Integer
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore VariableTypeNotDeclared
Sub Foo(arg1)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_Const_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    Const bar = 42
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "VariableTypeNotDeclaredInspection";
            var inspection = new VariableTypeNotDeclaredInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
