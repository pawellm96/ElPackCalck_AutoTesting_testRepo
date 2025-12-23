using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Newtonsoft.Json;

namespace ElPackCalc.Commands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ListElectricalCircuits : IExternalCommand
    {
        private ElectricalPackSettings _settings;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
  
            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            List<ElectricalSystem> circuits = new List<ElectricalSystem>();

            if (selectedIds.Count > 0)
            {
                foreach (ElementId id in selectedIds)
                {
                    Element el = doc.GetElement(id);
                    if (el == null) continue;

                    if (el is ElectricalSystem sys)
                        circuits.Add(sys);
                    else if (el is FamilyInstance fi && fi.MEPModel != null)
                    {
                        ISet<ElectricalSystem> systems = fi.MEPModel.GetElectricalSystems();
                        if (systems != null) circuits.AddRange(systems);
                    }
                }
            }
            else
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ElectricalSystem));
                circuits.AddRange(collector.Cast<ElectricalSystem>());
            }

            List<string> info = new List<string>();

            foreach (var circuit in circuits.OrderBy(c => c.Name))
            {
                string name = string.IsNullOrWhiteSpace(circuit.Name) ? "Без имени" : circuit.Name;

                int poles = GetPolesNumber(circuit);
                double voltage = GetCircuitVoltage(circuit);
                double current = GetCircuitCurrent(circuit, poles);
                double powerFactor = GetPowerFactor(circuit);
                double fullLength = GetCircuitLength(circuit);
                WireSize wireSize = GetWireSize(circuit);

                var calculationDetails = CalculateVoltageDropWithDetails(circuit, fullLength);
                double voltageDropPercent = calculationDetails.VoltageDropPercent;

                info.Add($"ЦЕПЬ: {name}");
                info.Add($"  ОСНОВНЫЕ ПАРАМЕТРЫ:");
                info.Add($"  • Полюсов: {poles}");
                info.Add($"  • Напряжение: {voltage:F0} В");
                info.Add($"  • Ток: {current:F2} А");
                info.Add($"  • Коэффициент мощности: {powerFactor:F3}");
                info.Add($"  • Длина цепи: {fullLength:F2} м");

                if (wireSize != null)
                {
                    info.Add($"  • Размер провода: {wireSize.ConductorSize}");
                    info.Add($"  • Сопротивление (Rc): {wireSize.Rc:F4} Ом/1000 футов");
                    info.Add($"  • Реактивное сопротивление (Xc): {wireSize.Xc:F4} Ом/1000 футов");
                }

                info.Add($"  РАСЧЕТ ПАДЕНИЯ НАПРЯЖЕНИЯ:");
                info.Add($"  • Коэффициент K: {calculationDetails.KFactor:F3}");
                info.Add($"  • Длина в футах: {calculationDetails.LengthInFeet:F2} футов");
                info.Add($"  • sin(φ): {calculationDetails.SinPhi:F4}");
                info.Add($"  • Активная составляющая: {calculationDetails.ResistancePart:F4}");
                info.Add($"  • Реактивная составляющая: {calculationDetails.ReactancePart:F4}");
                info.Add($"  • Полное сопротивление: {calculationDetails.ImpedancePart:F4}");
                info.Add($"  • Падение напряжения: {calculationDetails.VoltageDrop:F4} В");
                info.Add($"  • ПАДЕНИЕ НАПРЯЖЕНИЯ: {voltageDropPercent:F2}%");
                info.Add("");
            }

            TaskDialog.Show("Детальные результаты расчета", string.Join("\n", info));

            return Result.Succeeded;
        }

        private int GetPolesNumber(ElectricalSystem circuit)
        {
            Parameter p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_NUMBER_OF_POLES);
            return (p != null && p.HasValue) ? p.AsInteger() : -1;
        }

        private double GetCircuitVoltage(ElectricalSystem circuit)
        {
            Parameter p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_VOLTAGE);
            if (p != null && p.HasValue)
            {
                double voltageInternal = p.AsDouble();
                double voltageVolts = UnitUtils.ConvertFromInternalUnits(voltageInternal, UnitTypeId.Volts);
                return voltageVolts;
            }
            return -1;
        }

        private double GetCircuitCurrent(ElectricalSystem circuit, int poles)
        {
            if (poles == 1)
            {
                return GetApparentCurrent(circuit);
            }
            else if (poles == 2)
            {
                return GetCurrentForTwoPole(circuit);
            }
            else if (poles == 3)
            {
                return GetCurrentForThreePole(circuit);
            }
            else
            {
                return -1;
            }
        }

        private double GetCurrentForTwoPole(ElectricalSystem circuit)
        {
            double apparentCurrent = GetApparentCurrent(circuit);
            double currentA = GetPhaseCurrent(circuit, "A");
            double currentB = GetPhaseCurrent(circuit, "B");

            if (currentA <= 0 || currentB <= 0)
                return apparentCurrent;

            if (Math.Abs(currentA - currentB) < 0.001)
            {
                return apparentCurrent;
            }
            else if (currentA > currentB)
            {
                return currentA;
            }
            else
            {
                return currentB;
            }
        }

        private double GetCurrentForThreePole(ElectricalSystem circuit)
        {
            double apparentCurrent = GetApparentCurrent(circuit);
            double currentA = GetPhaseCurrent(circuit, "A");
            double currentB = GetPhaseCurrent(circuit, "B");
            double currentC = GetPhaseCurrent(circuit, "C");

            if (currentA <= 0 || currentB <= 0 || currentC <= 0)
                return apparentCurrent;

            if (Math.Abs(currentA - currentB) < 0.001 &&
                Math.Abs(currentB - currentC) < 0.001 &&
                Math.Abs(currentA - currentC) < 0.001)
            {
                return apparentCurrent;
            }

            if (currentA > currentB && currentA > currentC)
            {
                return currentA;
            }
            else if (currentB > currentA && currentB > currentC)
            {
                return currentB;
            }
            else if (currentC > currentA && currentC > currentB)
            {
                return currentC;
            }

            return apparentCurrent;
        }

        private double GetApparentCurrent(ElectricalSystem circuit)
        {
            Parameter p = circuit.LookupParameter("Apparent Current");
            if (p != null && p.HasValue)
                return p.AsDouble();

            return -1;
        }

        private double GetPhaseCurrent(ElectricalSystem circuit, string phase)
        {
            string paramName = $"Apparent Current Phase {phase}";
            Parameter p = circuit.LookupParameter(paramName);

            if (p != null && p.HasValue)
                return p.AsDouble();

            return -1;
        }

        private double GetPowerFactor(ElectricalSystem circuit)
        {
            Parameter p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_POWER_FACTOR);
            if (p != null && p.HasValue)
            {
                double pf = p.AsDouble();
                if (pf > 1)
                    return pf / 100.0;
                else
                    return pf;
            }

            return 0.85;
        }

        private double GetCircuitLength(ElectricalSystem circuit)
        {
            double length = -1;

            Parameter p = circuit.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_LENGTH_PARAM);
            if (p != null && p.HasValue)
            {
                length = p.AsDouble();
            }
            else
            {
                p = circuit.LookupParameter("Length");
                if (p != null && p.HasValue)
                {
                    length = p.AsDouble();
                }
                //else
                //{
                //    length = CalculateCircuitLength(circuit);
                //}
            }

            if (length <= 0) return -1;

            length = UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Meters);

            double reservePercent = GetWireReservePercent();
            return length * (1 + reservePercent / 100.0);
        }

        //private double CalculateCircuitLength(ElectricalSystem circuit)
        //{           
        //    ElementSet elementSet = circuit.Elements;

        //    double totalLength = 0;
        //    int count = 0;

        //    foreach (Element element in elementSet)
        //    {
        //        if (element != null)
        //        {
        //            Parameter lengthParam = element.LookupParameter("Length");
        //            if (lengthParam != null && lengthParam.HasValue)
        //            {
        //                totalLength += lengthParam.AsDouble();
        //                count++;
        //            }
        //        }
        //    }

        //    if (count > 0)
        //    {
        //        double averageLength = totalLength / count;
        //        averageLength = UnitUtils.ConvertFromInternalUnits(averageLength, UnitTypeId.Meters);
        //        return averageLength;
        //    }

        //    return 100.0; 
        //}

        private double GetWireReservePercent()
        {
            if (_settings?.TemplateTables?.Count > 0)
                return _settings.TemplateTables[0].WireReservePercent;
            return 0;
        }

        private WireSize GetWireSize(ElectricalSystem circuit)
        {
            Parameter wireSizeParam = circuit.LookupParameter("EP_Wire Size");
            string wireSizeName = null;

            if (wireSizeParam != null && wireSizeParam.HasValue)
            {
                wireSizeName = wireSizeParam.AsString();
            }
            else
            {
                Element circuitType = circuit.Document.GetElement(circuit.GetTypeId());
                if (circuitType != null)
                {
                    Parameter typeWireSizeParam = circuitType.LookupParameter("EP_Wire Size");
                    if (typeWireSizeParam != null && typeWireSizeParam.HasValue)
                    {
                        wireSizeName = typeWireSizeParam.AsString();
                    }
                }
            }

            return FindWireSizeInSettings(wireSizeName);
        }

        private WireSize FindWireSizeInSettings(string conductorSize)
        {
            if (string.IsNullOrEmpty(conductorSize))
                return GetDefaultWireSize();

            if (_settings?.WireSizeTables?.Count > 0)
            {
                var wireSizeTable = _settings.WireSizeTables[0];
                var wireSize = wireSizeTable.WireSizes?.FirstOrDefault(ws =>
                    ws.ConductorSize.Equals(conductorSize, StringComparison.OrdinalIgnoreCase));

                if (wireSize != null)
                    return wireSize;
            }

            return GetDefaultWireSize();
        }

        private WireSize GetDefaultWireSize()
        {
            return new WireSize { ConductorSize = "#14", Rc = 3.1, Xc = 0.073 };
        }

        private double GetKFactor(int poles)
        {
            if (poles == 1 || poles == 2)
                return 2.0;
            else if (poles == 3)
                return Math.Sqrt(3);
            else
                return 0;
        }

        private VoltageDropCalculation CalculateVoltageDropWithDetails(ElectricalSystem circuit, double lengthInMeters)
        {
            var result = new VoltageDropCalculation();

           
            result.Voltage = GetCircuitVoltage(circuit);
            result.Poles = GetPolesNumber(circuit);
            result.Current = GetCircuitCurrent(circuit, result.Poles);
            result.PowerFactor = GetPowerFactor(circuit);

            if (result.Voltage <= 0 || result.Current <= 0 || lengthInMeters <= 0)
            {
                result.VoltageDropPercent = -1;
                return result;
            }

            result.SinPhi = Math.Sqrt(1 - Math.Pow(result.PowerFactor, 2));

            result.WireSize = GetWireSize(circuit);
            if (result.WireSize == null)
            {
                result.VoltageDropPercent = -1;
                return result;
            }

            result.LengthInMeters = lengthInMeters;
            result.LengthInFeet = lengthInMeters * 3.28084;

            result.KFactor = GetKFactor(result.Poles);

            result.ResistancePart = result.WireSize.Rc * result.PowerFactor;
            result.ReactancePart = result.WireSize.Xc * result.SinPhi;
            result.ImpedancePart = result.ResistancePart + result.ReactancePart;
            result.VoltageDrop = (result.KFactor * result.Current * result.LengthInFeet * result.ImpedancePart) / 1000;

            result.VoltageDropPercent = (result.VoltageDrop / result.Voltage) * 100;

            return result;            
        }
    }

    public class VoltageDropCalculation
    {
        public double Voltage { get; set; }
        public int Poles { get; set; }
        public double Current { get; set; }
        public double PowerFactor { get; set; }
        public double SinPhi { get; set; }
        public WireSize WireSize { get; set; }
        public double LengthInMeters { get; set; }
        public double LengthInFeet { get; set; }
        public double KFactor { get; set; }
        public double ResistancePart { get; set; }
        public double ReactancePart { get; set; }
        public double ImpedancePart { get; set; }
        public double VoltageDrop { get; set; }
        public double VoltageDropPercent { get; set; }
    }

    public class ElectricalPackSettings
    {
        public List<WireSizeTable> WireSizeTables { get; set; }
        public List<TemplateTable> TemplateTables { get; set; }
    }

    public class WireSizeTable
    {
        public List<WireSize> WireSizes { get; set; }
    }

    public class WireSize
    {
        public string ConductorSize { get; set; }
        public double Rc { get; set; }
        public double Xc { get; set; }
    }

    public class TemplateTable
    {
        public double WireReservePercent { get; set; }
    }
}