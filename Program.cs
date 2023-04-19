namespace ListToXls {
    class Program {
        static void Main(string[] args) {
            List<Song> songs = new() {
                new() {
                    Title = "Super Slow",
                    Author = "Jaded Juice Riders"
                },
                new() {
                    Title = "Gutter Girl",
                    Author = "Hot Flash Heat Wave"
                },
                new() {
                    Title = "Satellite",
                    Author = "STRFKR"
                },
                new() {
                    Title = "Rose Pink Cadillac",
                    Author = "DOPE LEMON"
                },
                new() {
                    Title = "Rocky Said",
                    Author = "Dead Ghosts"
                },
                new() {
                    Title = "Lets Go Surfing",
                    Author = "The Drums"
                }
            };

            CreateExcel(songs, "songs");
        }

        private static string ListToHtmlTable<T>(List<T> items) {
            var properties = typeof(T).GetProperties();

            // Definisco lo stile della tabella
            string startHtml = @"
		        <html>
		        <head>
		        <style>
		        table, td, th {
		            border: 1px solid black;
		        }

		        table {
		            border-collapse: collapse;
		            width: 100%;
		        }

		        th {
		            height: 50px;
		            vertical-align: center;
		        }
		
		        td {
		            height: 50px;
		            vertical-align: top;
		        }
		        </style>
		        </head>	
		        <body>
	        ";

            // Costruisco la tabella
            string table = @"<table>";

            // Definisco l'header
            string row = @"<tr>";
            for (int j = 0; j < properties.Length; j++) {
                string col = string.Format(@"<th>{0}</th>", properties[j].Name);
                row += col;
            }
            row += @"</tr>";
            table += row;

            // Inserisco il contenuto principale
            for (int i = 0; i < items.Count; i++) {
                row = @"<tr>";
                for (int j = 0; j < properties.Length; j++) {
                    string col = string.Format(@"<td>{0}</td>", properties[j].GetValue(items[i]));
                    row += col;
                }
                row += @"</tr>";

                table += row;
            }
            table += @"</table>";

            string endHtml = @"</body></html>";

            return startHtml + table + endHtml;
        }

        private static void CreateExcel<T>(List<T> items, string filename) {
            Console.Write($"Creazione file {filename}.xls...");

            string content = ListToHtmlTable(items);
            Directory.CreateDirectory("Files");
            File.WriteAllText($"Files\\{filename}.xls", content);

            Console.WriteLine("Ok");
        }
    }
}