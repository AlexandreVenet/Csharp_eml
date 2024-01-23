// --------------------------------------------------------------
// On utilise le package https://www.nuget.org/packages/MSGReader
using MsgReader.Mime;
// --------------------------------------------------------------
// using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Csharp_eml
{
	internal class ExplorerEml
	{
		private string _cheminFichierEml = @"...\chemin\vers\message.eml";
		private string _dossierSortiePiecesJointes = @"...\chemin\vers\dossier\des\pieces\jointes";
		private string[] _extensionsFichierCherchees = new[] { ".XML", ".PDF" }; // Pouquoi pas minuscules ? Choix au dés ^^

		public void Demarrer()
		{
			// MSGReader réclame un FileInfo pour charger le fichier
			FileInfo infoFichier = new FileInfo(_cheminFichierEml);

			// Charger le fichier
			// On explicite le namespace pour lever une ambiguïté avec MsgReader.Outlook.Storage.Message
			MsgReader.Mime.Message eml = MsgReader.Mime.Message.Load(infoFichier);

			// On peut effectuer l'exploration. Voyons quelques-unes des données.
			// Attention, le fichier eml fourni peut ne pas contenir toutes les données d'un message valide.

			if (eml.Headers != null)
			{
				if (eml.Headers.To != null)
				{
					Console.WriteLine("Headers.To");
					foreach (MsgReader.Mime.Header.RfcMailAddress? recipient in eml.Headers.To)
					{
						Console.WriteLine($"\tAdress : " + recipient.Address);
						Console.WriteLine($"\tDisplayName : " + recipient.DisplayName);
						Console.WriteLine("\tHasValidMailAddress : " + recipient.HasValidMailAddress);
						Console.WriteLine("\tMailAddress : " + recipient.MailAddress);
					}
				}

				var subject = eml.Headers.Subject;
				Console.WriteLine("\nHeaders.Subject : " + subject);
			}

			// Corps en texte brut
			if (eml.TextBody != null)
			{
				var textBody = System.Text.Encoding.UTF8.GetString(eml.TextBody.Body);
				Console.WriteLine("\nTextBody : " + textBody);
			}

			// Corps en texte HTML
			if (eml.HtmlBody != null)
			{
				var htmlBody = System.Text.Encoding.UTF8.GetString(eml.HtmlBody.Body);
				Console.WriteLine("\nHtmlBody : " + htmlBody);
			}

			if (eml.Attachments != null)
			{
				Console.WriteLine("Attachments : " + eml.Attachments.Count);
				for (int i = 0; i < eml.Attachments.Count; i++)
				{
					MessagePart item = eml.Attachments[i];
					Console.WriteLine($"\tContentId : " + item.ContentId);
					Console.WriteLine($"\tFileName : " + item.FileName);
					Console.WriteLine($"\tContentDescription : " + item.ContentDescription);
					Console.WriteLine($"\tContentDisposition : " + item.ContentDisposition);
					Console.WriteLine($"\tContentType : " + item.ContentType);
					Console.WriteLine($"\tIsAttachment : " + item.IsAttachment);
					Console.WriteLine($"\tIsInline : " + item.IsInline);
					Console.WriteLine($"\tIsMultiPart : " + item.IsMultiPart);
					Console.WriteLine($"\tIsText : " + item.IsText);

					// Il existe une commande Save(), testons-la.

					// Chemin de sortie : dossier + nom de la pièce jointe (avec son extension).
					string pathFile = Path.Combine(_dossierSortiePiecesJointes, item.FileName);
					// Il faut un FileInfo
					FileInfo pathFileInfo = new(pathFile);
					// Cherchons les fichiers dont l'extension est ceci-cela
					if (Array.Exists(_extensionsFichierCherchees, (x) => pathFileInfo.Extension.ToUpper() == x))
					{
						item.Save(pathFileInfo);
					}
				}

				// Façon d'attendre que les fichiers soient bien récupérés
				while (true)
				{
					if (Directory.GetFiles(_dossierSortiePiecesJointes, "*.*", SearchOption.TopDirectoryOnly).Length == eml.Attachments.Count)
					{
						break;
					}
				}

				Console.WriteLine("\nToutes les pièces jointes ont été exportées.");
			}
		}
	}
}
