using TrackerCommon;

string? path;
do
{
    Console.WriteLine("Give me the path to an ods file:");
    path = Console.ReadLine();
} while (!File.Exists(path));

var test = new ODSTools();
test.DuplicateSpreadsheet(path);

Console.ReadLine();