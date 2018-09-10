using System;
using System.Collections.Generic;
using System.Linq;
using bumblebee.Types;
using CoreNodeModels;
using Dynamo.Graph.Nodes;
using Dynamo.Utilities;
using Newtonsoft.Json;
using ProtoCore.AST.AssociativeAST;

namespace bumblebee.UI
{
    [NodeName("Horizontal Align Types")]
    [NodeCategory("archilab_Bumblebee.Types")]
    [NodeDescription("Retrieve all available Horizontal Alignment types.")]
    [IsDesignScriptCompatible]
    public class HorizontalAlignTypesUi : DSDropDownBase
    {
        private const string OutputName = "Horizontal Align Type";
        private const string NoFamilyTypes = "No types were found.";
        public static HorizontalAlignTypes cscTypes = new HorizontalAlignTypes();

        public HorizontalAlignTypesUi() : base(OutputName) { }

        [JsonConstructor]
        public HorizontalAlignTypesUi(IEnumerable<PortModel> inPorts, IEnumerable<PortModel> outPorts) : base(OutputName, inPorts, outPorts) { }

        protected override SelectionState PopulateItemsCore(string currentSelection)
        {
            Items.Clear();

            var d = new Dictionary<string, int>(cscTypes.Types);

            if (d.Count == 0)
            {
                Items.Add(new DynamoDropDownItem(NoFamilyTypes, null));
                SelectedIndex = 0;
                return SelectionState.Done;
            }

            foreach (var pair in d)
            {
                Items.Add(new DynamoDropDownItem(pair.Key, pair.Value));
            }
            Items = Items.OrderBy(x => x.Name).ToObservableCollection();
            return SelectionState.Restore;
        }

        public override IEnumerable<AssociativeNode> BuildOutputAst(List<AssociativeNode> inputAstNodes)
        {
            if (Items.Count == 0 ||
                Items[0].Name == NoFamilyTypes ||
                SelectedIndex == -1)
            {
                return new[] { AstFactory.BuildAssignment(GetAstIdentifierForOutputIndex(0), AstFactory.BuildNullNode()) };
            }

            var args = new List<AssociativeNode>
            {
                AstFactory.BuildStringNode(Items[SelectedIndex].Name)
            };

            var func = new Func<string, int>(HorizontalAlignTypes.ByName);
            var functionCall = AstFactory.BuildFunctionCall(func, args);

            return new[] { AstFactory.BuildAssignment(GetAstIdentifierForOutputIndex(0), functionCall) };
        }
    }
}
