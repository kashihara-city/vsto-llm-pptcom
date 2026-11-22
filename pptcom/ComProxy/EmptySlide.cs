using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Pptcom.ComProxy
{
    // A slide can disappear only after the live checks below and a comparison with a slide created
    // by PowerPoint on the same layout. The comparison includes the notes and package relationships.
    internal static class EmptySlide
    {
        internal sealed class Candidate(object presentation, object slide, byte[] fingerprint, byte[]? savedFingerprint)
        {
            public object Presentation { get; } = presentation;
            public object Slide { get; } = slide;
            private byte[] Fingerprint { get; } = fingerprint;
            private byte[]? SavedFingerprint { get; } = savedFingerprint;

            public bool StillMatches(SlidePackage live, SlidePackage? saved)
            {
                bool sameFile = saved is null ? SavedFingerprint is null :
                    SavedFingerprint is not null && saved.Fingerprint.SequenceEqual(SavedFingerprint);
                return sameFile && live.Fingerprint.SequenceEqual(Fingerprint);
            }
        }

        public static Candidate? Find(object presentation, out bool untouched, out int placeholders, out string holds)
        {
            untouched = false;
            placeholders = 0;
            holds = "";
            try
            {
                // A legacy or foreign presentation stays without even querying its slide objects.
                _ = NativePath(presentation);
                object? slide = FindCandidate(presentation, out untouched, out placeholders, out holds);
                if (slide is null)
                {
                    return null;
                }

                bool stock = untouched;
                SlidePackage current = Capture(presentation, slide, stock, out SlidePackage? saved);
                if (stock)
                {
                    SlidePackage starter = Reference(presentation, slide, stock: true, current.Notes is not null);
                    stock = current.SameStockDesign(starter) && current.Matches(starter, slide, allowMoved: false);
                }

                bool empty = stock;
                if (!stock)
                {
                    bool withNotes = current.Notes is not null || saved?.Notes is not null;
                    empty = MatchesReference(Reference(presentation, slide, stock: false, withNotes));
                    if (!empty && (int)((dynamic)presentation).Designs.Count > 1)
                    {
                        // Reassigning a slide to a second master makes PowerPoint assign new default
                        // placeholder IDs and names. Reproduce the operation without ignoring names.
                        empty = MatchesReference(Reference(presentation, slide, stock: false, withNotes, reassign: true));
                    }

                    bool MatchesReference(SlidePackage blank)
                    {
                        return current.Matches(blank, slide, allowMoved: true) &&
                            (saved is null || current.AccountsFor(saved) || saved.Matches(blank, slide, allowMoved: false));
                    }
                }

                if (!empty)
                {
                    holds = "slide or notes data different from a new slide of its layout";
                }
                else if (SameSlide(presentation, slide))
                {
                    untouched = stock;
                    return new Candidate(presentation, slide, current.Fingerprint, saved?.Fingerprint);
                }
                else
                {
                    holds = "a change during inspection";
                }
            }
            catch (InvalidDataException ex)
            {
                holds = ex.Message;
            }
            catch (Exception ex) when (Guard.IsRecoverable(ex))
            {
                holds = CannotBeRead;
            }

            untouched = false;
            placeholders = 0;
            return null;
        }

        // Cheap live exclusions also catch data which Office cannot serialize. A candidate still needs the package proof.
        internal static object? FindCandidate(object presentation, out bool untouched, out int placeholders, out string holds)
        {
            bool asMade = false;
            int count = 0;
            string kept = "";
            object? found = Guard.Try<object?>(() =>
            {
                dynamic pres = presentation;
                if ((int)pres.Slides.Count != 1)
                {
                    return null;
                }

                object slide = pres.Slides.Item(1);
                (kept, count) = Guard.Try(() => Held(slide, presentation), (CannotBeRead, 0));
                if (kept.Length > 0)
                {
                    return null;
                }

                asMade = Guard.Try(() => MayBeStarter(presentation) && NeverSaved(presentation), fallback: false);
                return slide;
            }, fallback: null);
            untouched = found is not null && asMade;
            placeholders = found is null ? 0 : count;
            holds = found is null ? kept : "";
            return found;
        }

        // Whether `slide`, which Find found a while ago, is still the deck's one slide and still holds nothing. The
        // user was asked about that slide, and the deck can have changed while they answered: the slides are held
        // against the user's hand, not against another program. A slide put in its place is another slide, whatever
        // it holds (it has another SlideID, or the one asked about can no longer be read).
        public static bool IsStill(object presentation, Candidate candidate, bool requireUntouched)
        {
            return Guard.Try(() =>
            {
                if (!ReferenceEquals(presentation, candidate.Presentation) || !SameSlide(presentation, candidate.Slide) ||
                    (requireUntouched && ((string)((dynamic)presentation).Path).Length != 0) ||
                    Held(candidate.Slide, presentation).Holds.Length != 0)
                {
                    return false;
                }

                SlidePackage current = Capture(presentation, candidate.Slide, compareStock: false, out SlidePackage? saved);
                // On a new deck merely fetching BuiltInDocumentProperties can clear Saved, and
                // SaveCopyAs supplies last-save metadata. Neither is a second origin test.
                return candidate.StillMatches(current, saved) && SameSlide(presentation, candidate.Slide);
            }, fallback: false);
        }

        internal static bool SameSlide(object presentation, object slide)
        {
            return Guard.Try(() =>
            {
                dynamic slides = ((dynamic)presentation).Slides;
                return (int)slides.Count == 1 && (int)slides.Item(1).SlideID == (int)((dynamic)slide).SlideID;
            }, fallback: false);
        }

        private static readonly string[] OpenXmlExtensions = [".pptx", ".ppsx", ".potx"];

        private static string? NativePath(object presentation)
        {
            dynamic deck = presentation;
            if (((string)deck.Path).Length == 0)
            {
                return null;
            }

            string path = (string)deck.FullName;
            return OpenXmlExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase) ? path :
                throw new InvalidDataException("Inspection would require a format conversion.");
        }

        private static SlidePackage Capture(object presentation, object slide, bool compareStock, out SlidePackage? saved)
        {
            dynamic deck = presentation;
            if ((int)deck.PrintOptions.Ranges.Count != 0 || (int)deck.SlideShowSettings.NamedSlideShows.Count != 0 ||
                (int)deck.SlideShowSettings.RangeType != 1)
            {
                throw new InvalidDataException("A print range or slide show may depend on this slide.");
            }

            if ((bool)deck.HasVBProject || (bool)deck.Permission.Enabled || (int)deck.Signatures.Count != 0)
            {
                throw new InvalidDataException("The presentation has macros, rights management or signatures.");
            }

            int slideId = (int)((dynamic)slide).SlideID;
            saved = null;
            if (NativePath(presentation) is string path)
            {
                // A conversion (including PowerPoint reopening and reserializing its own file) can
                // discard an extension it does not understand. Inspect the persisted original too.
                using FileStream source = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                saved = SlidePackage.Read(source, slideId, compareStock: false);
            }

            return Snapshot(presentation, slideId, compareStock);
        }

        private static SlidePackage Reference(object presentation, object slide, bool stock, bool withNotes, bool reassign = false)
        {
            dynamic deck = ((dynamic)presentation).Application.Presentations.Add(0); // no window
            try
            {
                dynamic blank;
                if (stock)
                {
                    blank = deck.Slides.Add(1, 1); // the ordinary new title slide
                }
                else
                {
                    dynamic source = slide;
                    dynamic? initial = reassign ? deck.Slides.Add(1, 1) : null;
                    dynamic design = deck.Designs.Clone(source.Design);
                    dynamic layout = design.SlideMaster.CustomLayouts.Item((int)source.CustomLayout.Index);
                    if (initial is not null)
                    {
                        blank = initial;
                        blank.CustomLayout = layout;
                    }
                    else
                    {
                        blank = deck.Slides.AddSlide(1, layout);
                    }
                }

                if (withNotes)
                {
                    _ = blank.NotesPage.Shapes.Count;
                }

                return Snapshot((object)deck, (int)blank.SlideID, compareStock: stock);
            }
            finally
            {
                // Only this disposable, windowless presentation is closed.
                deck.Close();
            }
        }

        private static SlidePackage Snapshot(object presentation, int slideId, bool compareStock)
        {
            string copy = ComProxyEngine.TempPath("pptcom_empty_" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                ((dynamic)presentation).SaveCopyAs(copy, 24); // ppSaveAsOpenXMLPresentation
                using FileStream stream = File.OpenRead(copy);
                return SlidePackage.Read(stream, slideId, compareStock);
            }
            finally
            {
                File.Delete(copy);
            }
        }

        // The package written by PowerPoint. Follow the OPC relationships in both directions:
        // the slide and notes may own only their standard links, and no other part may point to
        // them. Unknown relationships/parts keep the slide, including newer comment formats.
        internal sealed class SlidePackage(XElement slideXml, XElement? notes, byte[] fingerprint, byte[] design, XElement slideSize, Uri layout, XElement options, bool section)
        {
            private const string OfficeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
            private const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/";
            private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
            private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
            private static readonly XNamespace R = OfficeRel.TrimEnd('/');
            private static readonly XNamespace P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
            private static readonly XNamespace A16 = "http://schemas.microsoft.com/office/drawing/2014/main";

            public XElement Slide { get; } = slideXml;
            public XElement? Notes { get; } = notes;
            public byte[] Fingerprint { get; } = fingerprint;
            private byte[] Design { get; } = design;
            private XElement SlideSize { get; } = slideSize;
            private Uri Layout { get; } = layout;
            private XElement Options { get; } = options;
            private bool HasSection { get; } = section;

            public bool AccountsFor(SlidePackage saved)
            {
                return XNode.DeepEquals(saved.Slide, Slide) && (saved.Notes is null || XNode.DeepEquals(saved.Notes, Notes)) &&
                    XNode.DeepEquals(saved.Options, Options);
            }

            public bool SameStockDesign(SlidePackage starter)
            {
                return !HasSection && Layout == starter.Layout && Design.SequenceEqual(starter.Design) && XNode.DeepEquals(SlideSize, starter.SlideSize);
            }

            public bool Matches(SlidePackage reference, object slide, bool allowMoved)
            {
                if (!XNode.DeepEquals(Options, reference.Options) || (Notes is not null && !XNode.DeepEquals(Notes, reference.Notes)))
                {
                    return false;
                }

                if (XNode.DeepEquals(Slide, reference.Slide))
                {
                    return true;
                }

                if (!allowMoved)
                {
                    return false;
                }

                // Ignore just the agreed exceptions: a slide name and the exact translation/size
                // of an empty frame that draws neither fill nor line. Never do this on the notes page.
                XElement comparable = new(Slide);
                XElement common = comparable.Element(P + "cSld");
                common.Attribute("name")?.Remove();
                XElement[] shapes = [.. common.Element(P + "spTree").Elements(P + "sp")];
                dynamic live = ((dynamic)slide).Shapes;
                if (shapes.Length != (int)live.Count)
                {
                    return false;
                }

                for (int i = 0; i < shapes.Length; i++)
                {
                    XElement shape = shapes[i];
                    XElement? transform = shape.Element(P + "spPr")?.Element(A + "xfrm");
                    if (transform is null || !PlacementOnly(transform))
                    {
                        continue;
                    }

                    dynamic item = live.Item(i + 1);
                    if ((int?)shape.Element(P + "nvSpPr")?.Element(P + "cNvPr")?.Attribute("id") == (int)item.Id &&
                        (int)item.Fill.Visible == 0 && (int)item.Line.Visible == 0 && (int)item.Shadow.Visible == 0 &&
                        (float)item.Glow.Radius == 0 && (int)item.Reflection.Type == 0 && (int)item.ThreeD.Visible == 0)
                    {
                        transform.Remove();
                    }
                }

                return XNode.DeepEquals(comparable, reference.Slide);
            }

            private static bool PlacementOnly(XElement transform)
            {
                XElement? offset = transform.Element(A + "off"), extent = transform.Element(A + "ext");
                if (offset is null || extent is null)
                {
                    return false;
                }

                XElement known = new(A + "xfrm",
                    new XElement(A + "off", new XAttribute("x", (long)offset.Attribute("x")), new XAttribute("y", (long)offset.Attribute("y"))),
                    new XElement(A + "ext", new XAttribute("cx", (long)extent.Attribute("cx")), new XAttribute("cy", (long)extent.Attribute("cy"))));
                return SameKnownMarkup(transform, known);
            }

            public static SlidePackage Read(Stream stream, int slideId, bool compareStock)
            {
                using Package package = Package.Open(stream, FileMode.Open, FileAccess.Read);
                Dictionary<Uri, string> roles = [];
                ReadRoles(package, roles);
                PackagePart presentation = Part("officeDocument");
                PackagePart slide = Part("slide");
                PackagePart? notes = roles.ContainsValue("notesSlide") ? Part("notesSlide") : null;
                XElement deck = Xml(presentation);
                XElement id = deck.Elements(P + "sldIdLst").Single().Elements().Single();
                string slideLink = (string)id.Attribute(R + "id");
                XElement onlySlide = new(P + "sldIdLst", new XElement(P + "sldId", new XAttribute("id", slideId), new XAttribute(R + "id", slideLink)));
                if (deck.Name != P + "presentation" || deck.Attributes().Any(static a => !a.IsNamespaceDeclaration && !DeckAttributes.Contains(a.Name)) ||
                    !SameKnownMarkup(id.Parent, onlySlide) ||
                    deck.Descendants().Attributes(R + "id").Any(a => a.Parent != id && a.Value == slideLink) ||
                    Target(presentation.GetRelationship(slideLink)) != slide.Uri ||
                    deck.Elements().Any(e => e != id.Parent && !KnownSetting(e, presentation, roles, slideId)))
                {
                    throw new InvalidDataException("The slide is referenced by an additional or unknown presentation setting.");
                }

                Uri layout = Target(slide.GetRelationshipsByType(OfficeRel + "slideLayout").Single());
                XElement content = Page(slide, "sld", Xml(package.GetPart(layout)));
                XElement? notesContent = notes is null ? null : Page(notes, "notes", layout: null);
                using MemoryStream fingerprint = new(), design = new();
                long budget = 64L * 1024 * 1024;
                foreach (PackagePart part in package.GetParts().OrderBy(static part => part.Uri.OriginalString, StringComparer.Ordinal))
                {
                    bool relationships = PackUriHelper.IsRelationshipPartUri(part.Uri);
                    Uri owner = relationships ? PackUriHelper.GetSourcePartUriFromRelationshipPartUri(part.Uri) : part.Uri;
                    string role = owner.OriginalString == "/" ? "/" : roles[owner];
                    if (VolatileParts.Contains(role))
                    {
                        continue;
                    }

                    bool shared = compareStock && DesignParts.Contains(role);
                    using Stream data = part.GetStream(FileMode.Open, FileAccess.Read);
                    budget -= data.Length;
                    if (budget < 0)
                    {
                        throw new InvalidDataException("The presentation is too large to compare during a request.");
                    }

                    byte[] digest;
                    using (SHA256 sha = SHA256.Create())
                    {
                        digest = sha.ComputeHash(data);
                    }

                    if (role == "core-properties" || role == "extended-properties")
                    {
                        // SaveCopyAs stamps a save time, revision and editing duration on the copy.
                        // Keep user metadata, including title, subject and description, in the recheck.
                        XElement metadata = Xml(part);
                        metadata.Elements().Where(static value => CopyStatistics.Contains(value.Name)).Remove();
                        digest = XmlHash(metadata);
                    }

                    WritePart(fingerprint, part, digest);
                    if (shared)
                    {
                        if (!relationships && (role == "slideLayout" || role == "slideMaster"))
                        {
                            XElement xml = Xml(part);
                            NormalizeCreationIds(xml);
                            foreach (XElement field in xml.Descendants(A + "fld"))
                            {
                                string? type = (string?)field.Attribute("type");
                                if (type == "slidenum" || type == "datetimeFigureOut")
                                {
                                    NormalizeGuid(field.Attribute("id"));
                                }
                            }

                            digest = XmlHash(xml);
                        }

                        if (role == "core-properties")
                        {
                            // Two genuine starters can of course have different creation times.
                            XElement metadata = Xml(part);
                            metadata.Elements().Where(static e => CopyStatistics.Contains(e.Name) || e.Name == XName.Get("created", "http://purl.org/dc/terms/")).Remove();
                            digest = XmlHash(metadata);
                        }

                        WritePart(design, part, digest);
                    }
                }

                return new SlidePackage(content, notesContent, fingerprint.ToArray(), design.ToArray(), deck.Elements(P + "sldSz").Single(), layout, Xml(Part("presProps")), deck.Descendants(P14 + "section").Any());

                PackagePart Part(string role)
                {
                    return package.GetPart(roles.Single(pair => pair.Value == role).Key);
                }
            }

            // PowerPoint can add an empty presentation-wide guide list simply by opening the file.
            // A named section also survives removal of its only slide. Match the complete structures
            // observed in the before/after packages, including when both share one extension list.
            private static bool KnownExtensions(XElement list, int slideId)
            {
                if (list.Name != P + "extLst")
                {
                    return false;
                }

                XElement expected = new(P + "extLst");
                HashSet<string> seen = new(StringComparer.Ordinal);
                foreach (XElement ext in list.Elements())
                {
                    const string Guides = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}", Sections = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
                    string uri = (string?)ext.Attribute("uri") ?? "";
                    if (!seen.Add(uri))
                    {
                        return false;
                    }

                    XElement contents;
                    if (uri == Guides)
                    {
                        contents = new XElement(XName.Get("sldGuideLst", "http://schemas.microsoft.com/office/powerpoint/2012/main"));
                    }
                    else if (uri == Sections)
                    {
                        XElement? section = ext.Element(P14 + "sectionLst")?.Element(P14 + "section");
                        string? id = (string?)section?.Attribute("id"), name = (string?)section?.Attribute("name");
                        if (name is null || !Guid.TryParseExact(id, "B", out _))
                        {
                            return false;
                        }

                        contents = new XElement(P14 + "sectionLst", new XElement(P14 + "section", new XAttribute("name", name), new XAttribute("id", id),
                            new XElement(P14 + "sldIdLst", new XElement(P14 + "sldId", new XAttribute("id", slideId)))));
                    }
                    else
                    {
                        return false;
                    }

                    expected.Add(new XElement(P + "ext", new XAttribute("uri", uri), contents));
                }

                return SameKnownMarkup(list, expected);
            }

            // Used only for fully specified structures containing numeric values or literal names.
            // Other XML, including markup-compatibility prefixes, is compared verbatim.
            private static bool SameKnownMarkup(XElement element, XElement expected)
            {
                XElement actual = new(element);
                actual.DescendantsAndSelf().Attributes().Where(static a => a.IsNamespaceDeclaration).Remove();
                return XNode.DeepEquals(actual, expected);
            }

            private static readonly HashSet<XName> DeckAttributes =
            [
                "saveSubsetFonts", "embedTrueTypeFonts", "autoCompressPictures", "showSpecialPlsOnTitleSld", "rtl",
                "firstSlideNum", "bookmarkIdSeed", "compatMode", "strictFirstAndLastChars", "conformance",
            ];

            // Inspect the contents of the presentation-wide lists too. A known outer tag does
            // not make a foreign nested slide reference safe, particularly in the original file.
            private static bool KnownSetting(XElement setting, PackagePart presentation, Dictionary<Uri, string> roles, int slideId)
            {
                if (setting.Name.Namespace != P)
                {
                    return false;
                }

                switch (setting.Name.LocalName)
                {
                    case "extLst":
                        return KnownExtensions(setting, slideId);
                    case "sldMasterIdLst":
                        return MasterList("sldMasterId", "slideMaster", numbered: true);
                    case "notesMasterIdLst":
                        return MasterList("notesMasterId", "notesMaster", numbered: false);
                    case "handoutMasterIdLst":
                        return MasterList("handoutMasterId", "handoutMaster", numbered: false);
                    case "sldSz":
                    case "notesSz":
                        XElement size = new(setting.Name, new XAttribute("cx", (long)setting.Attribute("cx")), new XAttribute("cy", (long)setting.Attribute("cy")));
                        if (setting.Name == P + "sldSz" && setting.Attribute("type") is XAttribute type)
                        {
                            size.Add(new XAttribute(type));
                        }

                        return SameKnownMarkup(setting, size);
                    case "defaultTextStyle":
                        return setting.DescendantsAndSelf().All(KnownTextStyleNode);
                    default:
                        return false;
                }

                bool MasterList(string entry, string role, bool numbered)
                {
                    XElement expected = new(setting.Name);
                    foreach (XElement item in setting.Elements())
                    {
                        string key = (string)item.Attribute(R + "id");
                        if (roles[Target(presentation.GetRelationship(key))] != role)
                        {
                            return false;
                        }

                        XElement known = new(P + entry);
                        if (numbered)
                        {
                            known.Add(new XAttribute("id", (uint)item.Attribute("id")));
                        }

                        known.Add(new XAttribute(R + "id", key));
                        expected.Add(known);
                    }

                    return SameKnownMarkup(setting, expected);
                }
            }

            private static bool KnownTextStyleNode(XElement node)
            {
                // These are text formatting properties, with no slide references. Keep a deck
                // with a richer default style until its additional markup can be accounted for.
                if (node.Nodes().Any(static n => n is not XElement && (n is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                {
                    return false;
                }

                if (node.Name == P + "defaultTextStyle")
                {
                    return node.Attributes().All(static attr => attr.IsNamespaceDeclaration);
                }

                return node.Name.Namespace == A && TextStyleElements.Contains(node.Name.LocalName) &&
                    node.Attributes().All(static attr => attr.IsNamespaceDeclaration || TextStyleAttributes.Contains(attr.Name));
            }

            private static readonly HashSet<string> TextStyleElements = new(StringComparer.Ordinal)
            {
                "defPPr", "lvl1pPr", "lvl2pPr", "lvl3pPr", "lvl4pPr", "lvl5pPr", "lvl6pPr", "lvl7pPr", "lvl8pPr", "lvl9pPr",
                "defRPr", "cs", "ea", "latin", "solidFill", "schemeClr",
            };

            private static readonly HashSet<XName> TextStyleAttributes =
            [
                "algn", "defTabSz", "eaLnBrk", "hangingPunct", "latinLnBrk", "marL", "rtl", "kern", "kumimoji", "lang", "sz", "typeface", "val",
            ];

            private static readonly HashSet<string> VolatileParts = new(StringComparer.Ordinal)
            {
                "thumbnail", "viewProps",
            };

            private static readonly HashSet<XName> CopyStatistics =
            [
                XName.Get("modified", "http://purl.org/dc/terms/"),
                XName.Get("revision", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"),
                XName.Get("lastModifiedBy", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"),
                XName.Get("TotalTime", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"),
            ];

            private static byte[] XmlHash(XElement xml)
            {
                using SHA256 sha = SHA256.Create();
                return sha.ComputeHash(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
            }

            private static readonly HashSet<string> DesignParts = new(StringComparer.Ordinal)
            {
                "officeDocument", "core-properties", "extended-properties", "tableStyles", "slideMaster", "slideLayout",
                "notesMaster", "handoutMaster", "theme", "themeOverride", "image",
            };

            private static void WritePart(Stream output, PackagePart part, byte[] digest)
            {
                using BinaryWriter writer = new(output, Encoding.UTF8, leaveOpen: true);
                writer.Write(part.Uri.OriginalString);
                writer.Write(part.ContentType);
                writer.Write(digest);
            }

            private static XElement Xml(PackagePart part)
            {
                using Stream data = part.GetStream(FileMode.Open, FileAccess.Read);
                using XmlReader reader = XmlReader.Create(data, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = 4 * 1024 * 1024,
                });
                return XElement.Load(reader, LoadOptions.PreserveWhitespace);
            }

            private static XElement Page(PackagePart part, string name, XElement? layout)
            {
                XElement xml = Xml(part);
                if (xml.Name != P + name || xml.Element(P + "cSld")?.Element(P + "spTree") is null)
                {
                    throw new InvalidDataException("The slide or notes page cannot be identified.");
                }

                NormalizeCreationIds(xml);
                if (layout is not null)
                {
                    foreach (XElement shape in xml.Element(P + "cSld").Element(P + "spTree").Elements(P + "sp"))
                    {
                        if ((string?)shape.Element(P + "nvSpPr")?.Element(P + "nvPr")?.Element(P + "ph")?.Attribute("type") == "pic")
                        {
                            RemoveEmptyPictureText(shape, layout);
                        }
                    }
                }
                else
                {
                    // PowerPoint assigns the automatic page-number field a fresh GUID on each page.
                    foreach (XElement field in xml.Descendants(A + "fld").Where(static f => (string?)f.Attribute("type") == "slidenum"))
                    {
                        NormalizeGuid(field.Attribute("id"));
                    }
                }

                return xml;
            }

            // On an empty picture placeholder PowerPoint sometimes materializes exactly this
            // empty paragraph with proofing languages. An attribute or node beyond this structure
            // (including an empty-looking text run, a hyperlink or an extension) remains evidence.
            private static void RemoveEmptyPictureText(XElement shape, XElement layout)
            {
                XElement? body = shape.Element(P + "txBody");
                XElement? end = body?.Element(A + "p")?.Element(A + "endParaRPr");
                if (body is null || end is null)
                {
                    return;
                }

                string? lang = (string?)end.Attribute("lang"), alternate = (string?)end.Attribute("altLang");
                if (lang is null || !layout.Descendants().Any(e => e.Name.Namespace == A &&
                    (string?)e.Attribute("lang") == lang && (string?)e.Attribute("altLang") == alternate))
                {
                    return;
                }

                XElement proofing = new(A + "endParaRPr");
                foreach (XAttribute attr in end.Attributes())
                {
                    if (attr.Name == "lang" || attr.Name == "altLang")
                    {
                        // A language must be one recognized by this installation.
                        if (attr.Value.Length == 0 || CultureInfo.GetCultureInfo(attr.Value).Name.Length == 0)
                        {
                            return;
                        }

                        proofing.Add(new XAttribute(attr));
                    }
                }

                XElement empty = new(P + "txBody", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"), new XElement(A + "p", proofing));
                if (SameKnownMarkup(body, empty))
                {
                    body.Remove();
                }
            }

            private static void NormalizeCreationIds(XElement xml)
            {
                foreach (XElement id in xml.Descendants(P14 + "creationId"))
                {
                    if (id.Parent?.Name == P + "ext" && (string?)id.Parent.Attribute("uri") == "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" &&
                        uint.TryParse((string?)id.Attribute("val"), NumberStyles.None, CultureInfo.InvariantCulture, out _))
                    {
                        id.SetAttributeValue("val", "0");
                    }
                }

                foreach (XElement id in xml.Descendants(A16 + "creationId"))
                {
                    if (id.Parent?.Name == A + "ext" && (string?)id.Parent.Attribute("uri") == "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}")
                    {
                        NormalizeGuid(id.Attribute("id"));
                    }
                }
            }

            private static void NormalizeGuid(XAttribute? attribute)
            {
                if (attribute is not null && Guid.TryParseExact(attribute.Value, "B", out _))
                {
                    attribute.Value = Guid.Empty.ToString("B");
                }
            }

            private static Uri Target(PackageRelationship link)
            {
                return PackUriHelper.ResolvePartUri(link.SourceUri, link.TargetUri);
            }

            private static void ReadRoles(Package package, Dictionary<Uri, string> roles)
            {
                Queue<(string Owner, PackageRelationshipCollection Links)> pending = new();
                pending.Enqueue(("/", package.GetRelationships()));
                while (pending.Count > 0)
                {
                    (string owner, PackageRelationshipCollection links) = pending.Dequeue();
                    HashSet<string> seen = new(StringComparer.Ordinal);
                    foreach (PackageRelationship link in links)
                    {
                        string role = RoleOf(link.RelationshipType);
                        if (link.TargetMode != TargetMode.Internal || !Children.TryGetValue(owner, out string[]? allowed) ||
                            !allowed.Contains(role, StringComparer.Ordinal))
                        {
                            throw new InvalidDataException("Additional or unknown package relationship.");
                        }

                        if (!seen.Add(role) && (owner == "slide" || owner == "notesSlide" || owner == "/" || role == "slide"))
                        {
                            throw new InvalidDataException("Ambiguous slide relationships.");
                        }

                        Uri target = Target(link);
                        if (roles.TryGetValue(target, out string? known))
                        {
                            if (role != known)
                            {
                                throw new InvalidDataException("Inconsistent part roles.");
                            }

                            continue;
                        }

                        PackagePart child = package.GetPart(target);
                        if (roles.Count >= 4096 || !KnownFormat(role, child.ContentType))
                        {
                            throw new InvalidDataException("Unsupported presentation part or size.");
                        }

                        roles.Add(target, role);
                        pending.Enqueue((role, child.GetRelationships()));
                    }

                    if ((owner == "slide" && !seen.Contains("slideLayout")) ||
                        (owner == "notesSlide" && (!seen.Contains("notesMaster") || !seen.Contains("slide"))) ||
                        (owner == "slideLayout" && !seen.Contains("slideMaster")))
                    {
                        throw new InvalidDataException("Incomplete slide relationships.");
                    }
                }
            }

            private static string RoleOf(string type)
            {
                return type switch
                {
                    PackageRel + "core-properties" => "core-properties",
                    PackageRel + "thumbnail" => "thumbnail",
                    OfficeRel + "thumbnail" or OfficeRel + "core-properties" => "",
                    _ => type.StartsWith(OfficeRel, StringComparison.Ordinal) ? type.Substring(OfficeRel.Length) : "",
                };
            }

            private const string DocumentType = "application/vnd.openxmlformats-officedocument.";
            private const string PresentationType = DocumentType + "presentationml.";

            private static bool KnownFormat(string role, string type)
            {
                return role switch
                {
                    "officeDocument" => type == PresentationType + "presentation.main+xml" || type == PresentationType + "template.main+xml" || type == PresentationType + "slideshow.main+xml",
                    "image" or "thumbnail" => type.StartsWith("image/", StringComparison.Ordinal),
                    "core-properties" => type == "application/vnd.openxmlformats-package.core-properties+xml",
                    "extended-properties" or "theme" or "themeOverride" => type == DocumentType + role + "+xml",
                    _ => type == PresentationType + role + "+xml",
                };
            }

            private static readonly Dictionary<string, string[]> Children = new(StringComparer.Ordinal)
            {
                ["/"] = ["officeDocument", "thumbnail", "core-properties", "extended-properties"],
                ["officeDocument"] = ["slide", "slideMaster", "notesMaster", "handoutMaster", "theme", "presProps", "viewProps", "tableStyles"],
                ["slide"] = ["slideLayout", "notesSlide"],
                ["notesSlide"] = ["slide", "notesMaster"],
                ["slideLayout"] = ["slideMaster", "image", "themeOverride"],
                ["slideMaster"] = ["slideLayout", "theme", "image"],
                ["notesMaster"] = ["theme", "image"],
                ["handoutMaster"] = ["theme", "image"],
                ["theme"] = ["image"],
                ["themeOverride"] = ["image"],
            };
        }

        private const string CannotBeRead = "something that cannot be read";

        // ---- 1. what the slide holds of its own ------------------------------------------------------------

        // What the slide holds, in words for the log ("" for nothing), and how many empty placeholders it has.
        private static (string Holds, int Placeholders) Held(object slide, object presentation)
        {
            foreach ((string what, Func<dynamic, dynamic, bool> there) in OnTheSlide)
            {
                if (there(slide, presentation))
                {
                    return (what, 0);
                }
            }

            if ((int)((dynamic)slide).HasNotesPage != 0 && OnTheNotesPage(slide, presentation) is string notes)
            {
                return (notes, 0);
            }

            dynamic shapes = ((dynamic)slide).Shapes;
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                object shape = shapes.Item(i);
                if (OnAShape(shape, slide) is string what)
                {
                    return (what, 0);
                }
            }

            return ("", count);
        }

        // What a slide can have set on itself, beside its shapes (the slide, then its presentation). Each is off, or
        // none, on a slide as PowerPoint makes it.
        private static readonly (string What, Func<dynamic, dynamic, bool> There)[] OnTheSlide =
        [
            ("a background of its own", static (s, _) => (int)s.FollowMasterBackground != -1),
            ("a comment", static (s, _) => (int)s.Comments.Count > 0),
            ("an animation", static (s, _) => (int)s.TimeLine.MainSequence.Count > 0 || (int)s.TimeLine.InteractiveSequences.Count > 0),
            ("a transition", static (s, _) => (int)s.SlideShowTransition.EntryEffect != 0),
            ("a time after which it advances", static (s, _) => (int)s.SlideShowTransition.AdvanceOnTime != 0),
            ("a sound for its transition", static (s, _) => (int)s.SlideShowTransition.SoundEffect.Type != 0),
            ("that it is hidden in the slide show", static (s, _) => (int)s.SlideShowTransition.Hidden != 0),
            ("a tag", static (s, _) => (int)s.Tags.Count > 0),
            ("a link", static (s, _) => (int)s.Hyperlinks.Count > 0),
            ("data a program attached to it", static (s, _) => (int)s.CustomerData.Count > 0),
            ("that the master's shapes are switched off on it", static (s, _) => (int)s.DisplayMasterShapes != -1),
            // PowerPoint takes such a layout away with the last slide that uses it; the layouts of its own themes are
            // nearly all kept, one that was added to a master is not
            ("a layout that goes when its last slide goes", static (s, _) => (int)s.CustomLayout.Preserved != -1),
            // and such a master likewise, with all its layouts and what is on them - when it is not the only one (the
            // one master of a presentation says "not preserved" too, and stays)
            ("a master that goes when its last slide goes", static (s, p) => (int)p.Designs.Count > 1 && (int)s.Design.Preserved != -1),
        ];

        // The notes page: text in its body placeholder; a shape that is not one of the page's placeholders (a text box
        // put there); a placeholder with something in it, hidden, with an alternative text, or with a look of its
        // own - not that of the notes master's placeholder of its kind (a picture set as the fill of the body shows on
        // the printed notes, with no text in it). The page's other placeholders are the slide's picture, the page
        // number and the like; a notes page has one of each kind, so the kind tells which of the master's is its own.
        private static string? OnTheNotesPage(object slide, object presentation)
        {
            dynamic shapes = ((dynamic)slide).NotesPage.Shapes;
            object ofMaster = ((dynamic)presentation).NotesMaster.Shapes;
            int count = (int)shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                object placeholder = shapes.Item(i);
                dynamic shape = placeholder;
                if ((int)shape.Type != MsoPlaceholder)
                {
                    return "a shape put on its notes page";
                }

                if ((int)shape.PlaceholderFormat.Type == PpPlaceholderBody && (int)shape.HasTextFrame != 0 && (int)shape.TextFrame.HasText != 0)
                {
                    return "notes";
                }

                if ((int)shape.PlaceholderFormat.ContainedType != MsoAutoShape || (int)shape.Visible == 0 ||
                    ((string?)shape.AlternativeText ?? "").Length > 0 || ((string?)shape.Title ?? "").Length > 0 ||
                    ShowsOfItsOwn(placeholder, ofMaster, atItsPlace: false))
                {
                    return "something set on a placeholder of its notes page";
                }
            }

            return null;
        }

        // What a shape on the slide says: anything but a placeholder that holds nothing, shows nothing of its own and
        // has nothing set on it.
        private static string? OnAShape(object placeholder, object slide)
        {
            dynamic shape = placeholder;
            if ((int)shape.Type != MsoPlaceholder)
            {
                return "a shape that is not a placeholder";
            }

            // PowerPoint says what a placeholder contains: an AutoShape while it is empty; a table, a picture, a chart
            // and the like once one was put in it - and then it has no text frame, so that "no text" alone would take
            // it for empty. An empty placeholder for a picture has no text frame either.
            if ((int)shape.PlaceholderFormat.ContainedType != MsoAutoShape)
            {
                return "a placeholder with a table, a picture or the like in it";
            }

            if ((int)shape.HasTextFrame != 0 && (int)shape.TextFrame.HasText != 0)
            {
                return "text in a placeholder";
            }

            foreach ((string what, Func<dynamic, bool> there) in OnAPlaceholder)
            {
                if (there(placeholder))
                {
                    return what;
                }
            }

            object ofLayout = ((dynamic)slide).CustomLayout.Shapes;
            return ShowsOfItsOwn(placeholder, ofLayout, atItsPlace: true) ? "a fill or a line of a placeholder's own" : null;
        }

        // What an empty placeholder can have set on itself. Each is off, or empty, on a slide as PowerPoint makes it.
        private static readonly (string What, Func<dynamic, bool> There)[] OnAPlaceholder =
        [
            ("a placeholder that was hidden", static h => (int)h.Visible == 0),
            ("an alternative text on a placeholder", static h => ((string?)h.AlternativeText ?? "").Length > 0 || ((string?)h.Title ?? "").Length > 0),
            ("an action on a placeholder", static h => (int)h.ActionSettings.Item(PpMouseClick).Action != PpActionNone || (int)h.ActionSettings.Item(PpMouseOver).Action != PpActionNone),
        ];

        // Whether the placeholder has a fill or a line that shows and is its own: not the look of its counterpart
        // among `ofTheMaster`, the shapes of its layout (or, for a notes page, of the notes master). Themes fill and
        // frame their placeholders, many of them with a picture - that is the layout's, like its shapes, and every
        // slide of the layout has it. The counterpart is the placeholder of the same kind, and (`atItsPlace`) at the
        // same place: a layout may have two of a kind, each with its own look, and one that was made to look like its
        // neighbour is not as its layout has it. A placeholder that shows and has no counterpart - it was moved, or
        // resized - counts as its own. The saved slide also has to match its reference.
        private static bool ShowsOfItsOwn(object placeholder, object ofTheMaster, bool atItsPlace)
        {
            dynamic shape = placeholder;
            if ((int)shape.Fill.Visible == 0 && (int)shape.Line.Visible == 0)
            {
                return false;
            }

            int kind = (int)shape.PlaceholderFormat.Type;
            string place = PlaceOf(placeholder), look = LookOf(placeholder);
            dynamic others = ofTheMaster;
            int count = (int)others.Count;
            for (int i = 1; i <= count; i++)
            {
                object its = others.Item(i);
                dynamic one = its;
                if ((int)one.Type == MsoPlaceholder && (int)one.PlaceholderFormat.Type == kind && (!atItsPlace || PlaceOf(its) == place))
                {
                    return LookOf(its) != look;
                }
            }

            return true;
        }

        // Where the shape is and how large, to a hundredth of a point.
        private static string PlaceOf(object of)
        {
            dynamic shape = of;
            return string.Join(" ", new[] { (float)shape.Left, (float)shape.Top, (float)shape.Width, (float)shape.Height }.Select(static v => v.ToString("0.##", CultureInfo.InvariantCulture)));
        }

        // A few inexpensive reasons to keep a visible frame before taking a copy. The absence of a
        // difference here is never enough to offer deletion. The package comparison checks the rest.
        private static string LookOf(object of)
        {
            dynamic h = of;
            return FormattableString.Invariant($"{(int)h.Fill.Visible}|{(int)h.Fill.Type}|{(int)h.Fill.ForeColor.RGB}|{(int)h.Line.Visible}|{(int)h.Line.ForeColor.RGB}");
        }

        private const int MsoPlaceholder = 14, MsoAutoShape = 1, PpPlaceholderBody = 2;
        private const int PpMouseClick = 1, PpMouseOver = 2, PpActionNone = 0;

        // ---- 2. whether nobody can have touched the presentation ---------------------------------------------

        // Saved is writable. These properties only make the deck eligible for comparison with a
        // stock presentation. An unreadable property cannot authorize the automatic path.
        public static bool MayBeStarter(object presentation)
        {
            return Guard.Try(() =>
            {
                dynamic pres = presentation;
                return ((string?)pres.Path)?.Length == 0 && (int)pres.Saved == -1;
            }, fallback: false);
        }

        // The built-in document properties by their place in the list (the same in every language of Office).
        private const int PropertyTemplate = 6, PropertyLastAuthor = 7, PropertyRevisionNumber = 8, PropertyCreationDate = 11, PropertyLastSaveTime = 12;

        // What reading a date that was never set fails with (E_FAIL).
        private const int NoValue = unchecked((int)0x80004005);

        // Whether nobody ever saved the presentation, in any form. One that PowerPoint made from nothing has no time of
        // a last save, and names no template, no last author and no revision; one made from a template, or opened as
        // an untitled copy of a file, brings those of the file. A time that was never set cannot be read, unlike one
        // that was. The time the presentation was made is always there, and is read first: a PowerPoint that answers
        // nothing at all must not be taken for one that says "never".
        private static bool NeverSaved(object presentation)
        {
            object properties = ((dynamic)presentation).BuiltInDocumentProperties;
            if (TextOf(properties, PropertyCreationDate).Length == 0 || TextOf(properties, PropertyTemplate).Length > 0 ||
                TextOf(properties, PropertyLastAuthor).Length > 0 || TextOf(properties, PropertyRevisionNumber).Length > 0)
            {
                return false;
            }

            try
            {
                return TextOf(properties, PropertyLastSaveTime).Length == 0;
            }
            catch (COMException ex) when (ex.HResult == NoValue)
            {
                return true;
            }
        }

        private static string TextOf(object properties, int index)
        {
            object? value = ((dynamic)properties).Item(index).Value;
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        }
    }
}
