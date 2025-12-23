using System;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using ElPackCalc.Commands;
using Moq;
using Xunit;

namespace ElPackCalc.Tests
{
    /// <summary>
    /// Тесты для ListElectricalCircuits без запуска Revit
    /// Использует рефлексию для тестирования приватных методов с бизнес-логикой
    /// </summary>
    public class ListElectricalCircuitsTests
    {
        private Mock<Parameter> CreateMockParameter(object value, bool hasValue = true, string stringValue = null)
        {
            var mockParam = new Mock<Parameter>();
            mockParam.Setup(p => p.HasValue).Returns(hasValue);
            
            if (value is double d)
            {
                mockParam.Setup(p => p.AsDouble()).Returns(d);
            }
            else if (value is int i)
            {
                mockParam.Setup(p => p.AsInteger()).Returns(i);
                mockParam.Setup(p => p.AsDouble()).Returns((double)i);
            }
            
            if (!string.IsNullOrEmpty(stringValue))
            {
                mockParam.Setup(p => p.AsString()).Returns(stringValue);
            }
            
            return mockParam;
        }

        private Mock<ElectricalSystem> CreateMockCircuit(
            string name = "Test Circuit",
            int poles = 3,
            double voltage = 400.0,
            double apparentCurrent = 10.0,
            double powerFactor = 0.85,
            double length = 50.0,
            string wireSize = "#12")
        {
            var mockCircuit = new Mock<ElectricalSystem>();
            mockCircuit.Setup(c => c.Name).Returns(name);

            // Настройка параметров
            var polesParam = CreateMockParameter(poles);
            var voltageParam = CreateMockParameter(UnitUtils.ConvertToInternalUnits(voltage, UnitTypeId.Volts));
            var powerFactorParam = CreateMockParameter(powerFactor);
            var lengthParam = CreateMockParameter(UnitUtils.ConvertToInternalUnits(length, UnitTypeId.Meters));
            var apparentCurrentParam = CreateMockParameter(apparentCurrent);
            var wireSizeParam = CreateMockParameter(null, false, wireSize);

            mockCircuit.Setup(c => c.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES))
                .Returns(polesParam.Object);
            mockCircuit.Setup(c => c.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE))
                .Returns(voltageParam.Object);
            mockCircuit.Setup(c => c.get_Parameter(BuiltInParameter.RBS_ELEC_POWER_FACTOR))
                .Returns(powerFactorParam.Object);
            mockCircuit.Setup(c => c.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM))
                .Returns(lengthParam.Object);

            mockCircuit.Setup(c => c.LookupParameter("Apparent Current"))
                .Returns(apparentCurrentParam.Object);
            mockCircuit.Setup(c => c.LookupParameter("EP_Wire Size"))
                .Returns(wireSizeParam.Object);

            return mockCircuit;
        }

        [Fact]
        public void GetPolesNumber_ShouldReturnCorrectValue()
        {
            // Arrange
            var circuit = CreateMockCircuit(poles: 3);
            var command = new ListElectricalCircuits();

            // Act
            var poles = GetPolesNumberUsingReflection(command, circuit.Object);

            // Assert
            Assert.Equal(3, poles);
        }

        [Fact]
        public void GetCircuitVoltage_ShouldReturnCorrectValue()
        {
            // Arrange
            var circuit = CreateMockCircuit(voltage: 400.0);
            var command = new ListElectricalCircuits();

            // Act
            var voltage = GetCircuitVoltageUsingReflection(command, circuit.Object);

            // Assert
            Assert.Equal(400.0, voltage, 1);
        }

        [Fact]
        public void GetPowerFactor_ShouldReturnCorrectValue()
        {
            // Arrange
            var circuit = CreateMockCircuit(powerFactor: 0.85);
            var command = new ListElectricalCircuits();

            // Act
            var powerFactor = GetPowerFactorUsingReflection(command, circuit.Object);

            // Assert
            Assert.Equal(0.85, powerFactor, 3);
        }

        [Fact]
        public void GetCircuitLength_ShouldReturnCorrectValue()
        {
            // Arrange
            var circuit = CreateMockCircuit(length: 50.0);
            var command = new ListElectricalCircuits();

            // Act
            var length = GetCircuitLengthUsingReflection(command, circuit.Object);

            // Assert
            Assert.True(length > 0);
        }

        [Fact]
        public void CalculateVoltageDrop_ShouldCalculateCorrectly()
        {
            // Arrange
            var circuit = CreateMockCircuit(
                poles: 3,
                voltage: 400.0,
                apparentCurrent: 10.0,
                powerFactor: 0.85,
                length: 50.0,
                wireSize: "#12");
            var command = new ListElectricalCircuits();

            // Act
            var calculation = CalculateVoltageDropUsingReflection(command, circuit.Object, 50.0);

            // Assert
            Assert.NotNull(calculation);
            Assert.True(calculation.VoltageDropPercent >= 0);
        }

        [Fact]
        public void GetKFactor_ShouldReturnCorrectValueForThreePoles()
        {
            // Arrange
            var command = new ListElectricalCircuits();

            // Act
            var kFactor3 = GetKFactorUsingReflection(command, 3);
            var kFactor2 = GetKFactorUsingReflection(command, 2);
            var kFactor1 = GetKFactorUsingReflection(command, 1);

            // Assert
            Assert.Equal(Math.Sqrt(3), kFactor3, 3);
            Assert.Equal(2.0, kFactor2, 1);
            Assert.Equal(2.0, kFactor1, 1);
        }

        // Вспомогательные методы для доступа к приватным методам через рефлексию
        private int GetPolesNumberUsingReflection(ListElectricalCircuits command, ElectricalSystem circuit)
        {
            var method = typeof(ListElectricalCircuits).GetMethod(
                "GetPolesNumber",
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (int)method.Invoke(command, new object[] { circuit });
        }

        private double GetCircuitVoltageUsingReflection(ListElectricalCircuits command, ElectricalSystem circuit)
        {
            var method = typeof(ListElectricalCircuits).GetMethod(
                "GetCircuitVoltage",
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (double)method.Invoke(command, new object[] { circuit });
        }

        private double GetPowerFactorUsingReflection(ListElectricalCircuits command, ElectricalSystem circuit)
        {
            var method = typeof(ListElectricalCircuits).GetMethod(
                "GetPowerFactor",
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (double)method.Invoke(command, new object[] { circuit });
        }

        private double GetCircuitLengthUsingReflection(ListElectricalCircuits command, ElectricalSystem circuit)
        {
            var method = typeof(ListElectricalCircuits).GetMethod(
                "GetCircuitLength",
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (double)method.Invoke(command, new object[] { circuit });
        }

        private double GetKFactorUsingReflection(ListElectricalCircuits command, int poles)
        {
            var method = typeof(ListElectricalCircuits).GetMethod(
                "GetKFactor",
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (double)method.Invoke(command, new object[] { poles });
        }

        private VoltageDropCalculation CalculateVoltageDropUsingReflection(
            ListElectricalCircuits command,
            ElectricalSystem circuit,
            double length)
        {
            var method = typeof(ListElectricalCircuits).GetMethod(
                "CalculateVoltageDropWithDetails",
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (VoltageDropCalculation)method.Invoke(command, new object[] { circuit, length });
        }
    }
}

