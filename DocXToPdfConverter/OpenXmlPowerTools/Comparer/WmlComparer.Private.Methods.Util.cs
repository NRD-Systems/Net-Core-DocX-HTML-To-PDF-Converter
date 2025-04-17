// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static partial class WmlComparer
    {
        private static XElement MoveRelatedPartsToDestination(
            IPackagePart partOfDeletedContent,
            IPackagePart partInNewDocument,
            XElement contentElement)
        {
            List<XElement> elementsToUpdate = contentElement
                .Descendants()
                .Where(d => d.Attributes().Any(a => ComparisonUnitWord.RelationshipAttributeNames.Contains(a.Name)))
                .ToList();

            foreach (XElement element in elementsToUpdate)
            {
                List<XAttribute> attributesToUpdate = element
                    .Attributes()
                    .Where(a => ComparisonUnitWord.RelationshipAttributeNames.Contains(a.Name))
                    .ToList();

                foreach (XAttribute att in attributesToUpdate)
                {
                    var rId = (string) att;

                    IPackageRelationship relationshipForDeletedPart = partOfDeletedContent.Relationships.Single(p => p.Id == rId);

                    Uri targetUri = PackUriHelper
                        .ResolvePartUri(
                            new Uri(partOfDeletedContent.Uri.ToString(), UriKind.Relative),
                            relationshipForDeletedPart.TargetUri);

                    IPackagePart relatedPackagePart = partOfDeletedContent.Package.GetPart(targetUri);
                    string[] uriSplit = relatedPackagePart.Uri.ToString().Split('/');
                    string[] last = uriSplit[uriSplit.Length - 1].Split('.');
                    string uriString;
                    if (last.Length == 2)
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                                    "P" + Guid.NewGuid().ToString().Replace("-", "") + "." + last[1];
                    }
                    else
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                                    "P" + Guid.NewGuid().ToString().Replace("-", "");
                    }

                    Uri uri = relatedPackagePart.Uri.IsAbsoluteUri
                        ? new Uri(uriString, UriKind.Absolute)
                        : new Uri(uriString, UriKind.Relative);

                    IPackagePart newPart = partInNewDocument.Package.CreatePart(uri, relatedPackagePart.ContentType, CompressionOption.Normal);

                    // ReSharper disable once PossibleNullReferenceException
                    using (Stream oldPartStream = relatedPackagePart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    using (Stream newPartStream = newPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                    {
                        FileUtils.CopyStream(oldPartStream, newPartStream);
                    }

                    string newRid = "R" + Guid.NewGuid().ToString().Replace("-", "");
                    partInNewDocument.Relationships.Create(newPart.Uri, TargetMode.Internal,
                        relationshipForDeletedPart.RelationshipType, newRid);
                    att.Value = newRid;

                    if (newPart.ContentType.EndsWith("xml"))
                    {
                        XDocument newPartXDoc;
                        using (Stream stream = newPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                        {
                            newPartXDoc = XDocument.Load(stream);
                            MoveRelatedPartsToDestination(relatedPackagePart, newPart, newPartXDoc.Root);
                        }

                        using (Stream stream = newPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                            newPartXDoc.Save(stream);
                    }
                }
            }

            return contentElement;
        }

        private static XAttribute GetXmlSpaceAttribute(string textOfTextElement)
        {
            if (char.IsWhiteSpace(textOfTextElement[0]) ||
                char.IsWhiteSpace(textOfTextElement[textOfTextElement.Length - 1]))
                return new XAttribute(XNamespace.Xml + "space", "preserve");

            return null;
        }
    }
}
