#if !NETSTANDARD2_0
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeDevPnP.Core.Tests.Framework.Providers.Extensibility
{
    public class SecureXml
    {
        private static Rijndael rKey = new RijndaelManaged();

        static SecureXml()
        {
            rKey.GenerateKey();
            rKey.GenerateIV();
        }

        private SecureXml() { }

        public static void SignXmlDocument(Stream sourceXmlFile,
            Stream destinationXmlFile, X509Certificate2 certificate)
        {
            // Carico il documento XML
            XmlDocument doc = new XmlDocument();
            doc.Load(sourceXmlFile);

            // Preparo un DOMDocument che conterrà il risultato
            XmlDocument outputDocument = new XmlDocument();

            // Recupero un riferimento all'intero contenuto del documento XML
            XmlNodeList elementsToSign = doc.SelectNodes(String.Format("/{0}", doc.DocumentElement.Name));

            // Costruisco la firma
            SignedXml signedXml = new SignedXml();
            System.Security.Cryptography.Xml.DataObject dataSignature =
                new System.Security.Cryptography.Xml.DataObject();
            dataSignature.Data = elementsToSign;
            dataSignature.Id = doc.DocumentElement.Name;
            signedXml.AddObject(dataSignature);
            Reference reference = new Reference();
            reference.Uri = String.Format("#{0}", dataSignature.Id);
            signedXml.AddReference(reference);

            if ((certificate != null) && (certificate.HasPrivateKey))
            {
                signedXml.SigningKey = certificate.PrivateKey;

                KeyInfo keyInfo = new KeyInfo();
                keyInfo.AddClause(new KeyInfoX509Data(certificate));
                signedXml.KeyInfo = keyInfo;
                signedXml.ComputeSignature();

                // Aggiungo la firma al nuovo documento di output
                outputDocument.AppendChild(
                    outputDocument.ImportNode(signedXml.GetXml(), true));

                outputDocument.Save(destinationXmlFile);
            }
        }

        public static Boolean CheckSignedXmlDocument(Stream sourceXmlFile)
        {
            // Carico il documento XML
            XmlDocument doc = new XmlDocument();
            doc.Load(sourceXmlFile);

            // Verifico la firma
            SignedXml sigs = new SignedXml(doc);
            XmlNodeList sigElems = doc.GetElementsByTagName("Signature");
            sigs.LoadXml((XmlElement)sigElems[0]);
            return (sigs.CheckSignature());
        }

        public static void EncryptXmlDocument(Stream sourceXmlFile,
            Stream destinationXmlFile,
            String xpathNodeToEncrypt,
            Dictionary<String, String> namespaces,
            X509Certificate2 certificate)
        {
            // Carico il documento XML
            XmlDocument doc = new XmlDocument();
            doc.Load(sourceXmlFile);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(doc.NameTable);
            if (namespaces != null)
            {
                foreach (var prefix in namespaces.Keys)
                {
                    namespaceManager.AddNamespace(prefix, namespaces[prefix]);
                }
            }

            // Estraggo l'elemento da cifrare
            XmlElement elementToEncrypt = doc.SelectSingleNode(xpathNodeToEncrypt, namespaceManager) as XmlElement;

            // Cifro il nodo di tipo elemento
            EncryptedXml enc = new EncryptedXml(doc);
            EncryptedData ed = enc.Encrypt(elementToEncrypt, certificate);

            // Lo sostituisco all'elemento originale
            EncryptedXml.ReplaceElement(elementToEncrypt,
                ed, false);

            // Salvo il risultato
            doc.Save(destinationXmlFile);
        }

        public static void DecryptXmlDocument(
            Stream sourceXmlFile,
            Stream destinationXmlFile,
            X509Certificate2 certificate)
        {
            // Apro il documento
            XmlDocument doc = new XmlDocument();
            doc.Load(sourceXmlFile);

            // Decifro il nodo di tipo elemento
            EncryptedXml enc = new EncryptedXml(doc);
            enc.DecryptDocument();

            // Salvo il risultato
            doc.Save(destinationXmlFile);
        }
    }
}
#endif