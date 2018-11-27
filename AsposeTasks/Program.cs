using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Tasks;

namespace AsposeTasks
{
    internal class Program
    {
        #region Static Methods

        public static void Main(string[] args)
        {
            var project = new Project("Data/Project1.mpp");

            ExtendedAttributeDefinitionCollection attributeCollection = project.ExtendedAttributes;

            foreach (ExtendedAttributeDefinition extendedAttributeDefinition in attributeCollection)
            {
                Console.WriteLine($"Extended Attribute {extendedAttributeDefinition.FieldName} with alias {extendedAttributeDefinition.Alias} and type {extendedAttributeDefinition.ElementType.ToString()}");
            }

            ExtendedAttributeDefinition extendedAttribute = project.ExtendedAttributes.GetById((int)ExtendedAttributeTask.Number1);

            if (extendedAttribute.ValueList != null && extendedAttribute.ValueList.Count != 0)
            {
                foreach (Value value in extendedAttribute.ValueList)
                {
                    Console.WriteLine($"Value can be {value.StringValue} with description {value.Description}");
                }
            }

            ExtendedAttributeDefinition extendedAttributeDate = project.ExtendedAttributes.GetById((int)ExtendedAttributeTask.Date1);

            string stringFormula = extendedAttributeDate.Formula;

            Console.WriteLine($"Date Formula is {stringFormula}");

            List<Task> taskList = project.SelectAllChildTasks().ToList();

            var valueRetrivalDictionary = new Dictionary<string, Func<ExtendedAttribute, string>>
            {
                { "Text", s => s.TextValue },
                { "Cost", s => s.NumericValue.ToString() },
                { "Number", s => s.NumericValue.ToString() },
                { "Date", s => s.DateValue.ToString() }
            };

            foreach (Task task in taskList)
            {
                foreach (ExtendedAttribute taskExtendedAttribute in task.ExtendedAttributes)
                {
                    string taskValue = string.Empty;

                    string fieldStringType =
                        taskExtendedAttribute.AttributeDefinition.FieldName.Substring(0,
                                                                                      taskExtendedAttribute.AttributeDefinition.FieldName
                                                                                          .Length - 1);
                    valueRetrivalDictionary.TryGetValue(fieldStringType, out Func<ExtendedAttribute, string> valueGetter);

                    if (valueGetter != null)
                    {
                        taskValue = valueGetter(taskExtendedAttribute);
                    }

                    Console.WriteLine($"Task {task} extended attribute {taskExtendedAttribute.AttributeDefinition.Alias} value set to {taskValue}");
                }
            }

            foreach (Resource projectResource in project.Resources)
            {
                foreach (ExtendedAttribute projectResourceExtendedAttribute in projectResource.ExtendedAttributes)
                {
                    string taskValue = string.Empty;

                    string fieldStringType =
                        projectResourceExtendedAttribute.AttributeDefinition.FieldName.Substring(0,
                                                                                                 projectResourceExtendedAttribute
                                                                                                     .AttributeDefinition.FieldName.Length
                                                                                                 - 1);
                    valueRetrivalDictionary.TryGetValue(fieldStringType, out Func<ExtendedAttribute, string> valueGetter);

                    if (valueGetter != null)
                    {
                        taskValue = valueGetter(projectResourceExtendedAttribute);
                    }

                    Console.WriteLine($"Resource {projectResource} extended attribute {projectResourceExtendedAttribute.AttributeDefinition.Alias} value set to {taskValue}");
                }
            }
        }

        #endregion
    }
}
