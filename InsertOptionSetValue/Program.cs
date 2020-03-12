using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.IO;

namespace InsertOptionSetValue
{
    class Program
    {
        static void Main(string[] args)
        {
            string optionSetName = "";
            string csvFile = "";
            string connectionString = "";
            int optionValueStart = 0;

            try {
                if (args.Length != 4)
                    throw new Exception($"ERROR: Incorrect number of arguments provided.\nExpected 3 arguments and received {args.Length}");

                optionSetName = args[0];

                if (!int.TryParse(args[1], out optionValueStart))
                    throw new Exception("OptionValueStart must be an Integer");

                csvFile = args[2];

                connectionString = args[3];
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine();
                Console.WriteLine("Usage: InsertOptionSetValue.exe [OptionSetLogicalName] [OptionValueStart] [File Path] [Organisation Connection String]");
                Console.ReadKey();
                return;
            }

            //connect to CDS
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            IOrganizationService orgService = (IOrganizationService)conn.OrganizationServiceProxy;

            int value = optionValueStart;

            //open csv file
            using (var reader = new StreamReader(csvFile))
            {
                //read csv file line by line
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');

                    // Create a request.
                    InsertOptionValueRequest insertOptionValueRequest =
                        new InsertOptionValueRequest
                        {
                            OptionSetName = optionSetName,
                            Label = new Label(values[0], 1033),
                            Value = value
                        };

                    // Execute the request.
                    int insertOptionValue = ((InsertOptionValueResponse)orgService.Execute(insertOptionValueRequest)).NewOptionValue;

                    value++;

                    Console.WriteLine("Created {0} with the value of {1}.", insertOptionValueRequest.Label.LocalizedLabels[0].Label, insertOptionValue);
                }
            }
            Console.WriteLine("\n******** COMPLETE ********");
            Console.WriteLine($"Inserted {value - optionValueStart} values into {optionSetName}");
        }
    }
}
