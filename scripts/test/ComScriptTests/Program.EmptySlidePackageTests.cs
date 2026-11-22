using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using Pptcom.ComProxy;
using static System.FormattableString;

namespace ComScriptTests
{
    internal static partial class Program
    {
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        private static readonly XNamespace PresentationMl = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace DrawingMl = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Office2010 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

        // Minimal, in-memory OPC presentations. The file/order numbers deliberately differ from
        // the slide ID. Every test passes through the same relationship reader as the add-in.
        private static EmptySlide.SlidePackage PackageWith(Action<PackagePart, PackagePart>? link = null,
            Action<XElement>? onSlide = null, Action<XElement>? onNotes = null, Action<XElement>? onDeck = null,
            Action<XElement>? onMaster = null, Action<XElement>? onOptions = null, int expectedId = 259)
        {
            XNamespace p = PresentationMl, a = DrawingMl, r = Rel.TrimEnd('/');
            using MemoryStream bytes = new();
            using (Package package = Package.Open(bytes, FileMode.Create, FileAccess.ReadWrite))
            {
                XElement deck = new(p + "presentation", new XElement(p + "sldMasterIdLst",
                    new XElement(p + "sldMasterId", new XAttribute("id", 2147483648U), new XAttribute(r + "id", "master"))),
                    new XElement(p + "sldIdLst", new XElement(p + "sldId", new XAttribute("id", 259), new XAttribute(r + "id", "page"))),
                    new XElement(p + "sldSz", new XAttribute("cx", 100), new XAttribute("cy", 80)));
                XElement frame = new(p + "sp", new XElement(p + "nvSpPr", new XElement(p + "cNvPr", new XAttribute("id", 2), new XAttribute("name", "Title 1")),
                    new XElement(p + "nvPr", new XElement(p + "ph", new XAttribute("type", "title")))), new XElement(p + "spPr"),
                    new XElement(p + "txBody", new XElement(a + "p")));
                XElement page = new(p + "sld", new XElement(p + "cSld", new XElement(p + "spTree", frame),
                    new XElement(p + "extLst", new XElement(p + "ext", new XAttribute("uri", "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"),
                        new XElement(Office2010 + "creationId", new XAttribute("val", "100"))))));
                XElement notes = new(p + "notes", new XElement(p + "cSld", new XElement(p + "spTree", new XElement(frame))));
                XElement masterXml = new(p + "sldMaster"), options = new(p + "presentationPr");
                onSlide?.Invoke(page);
                onNotes?.Invoke(notes);
                onDeck?.Invoke(deck);
                onMaster?.Invoke(masterXml);
                onOptions?.Invoke(options);

                PackagePart main = Part("/document.xml", deck, "presentation.main"),
                    slide = Part("/pages/7.xml", page, "slide"), note = Part("/notes/9.xml", notes, "notesSlide"),
                    layout = Part("/layouts/3.xml", new XElement(p + "sldLayout", new XAttribute("preserve", 1),
                        new XElement(a + "endParaRPr", new XAttribute("lang", "ja-JP"), new XAttribute("altLang", "en-US"))), "slideLayout"),
                    master = Part("/masters/1.xml", masterXml, "slideMaster"),
                    notesMaster = Part("/masters/2.xml", new XElement(p + "notesMaster"), "notesMaster");
                package.CreateRelationship(main.Uri, TargetMode.Internal, Rel + "officeDocument", "officeDocument");
                main.CreateRelationship(slide.Uri, TargetMode.Internal, Rel + "slide", "page");
                main.CreateRelationship(master.Uri, TargetMode.Internal, Rel + "slideMaster", "master");
                main.CreateRelationship(Part("/options.xml", options, "presProps").Uri, TargetMode.Internal, Rel + "presProps", "presProps");
                master.CreateRelationship(layout.Uri, TargetMode.Internal, Rel + "slideLayout", "slideLayout");
                layout.CreateRelationship(master.Uri, TargetMode.Internal, Rel + "slideMaster", "slideMaster");
                slide.CreateRelationship(layout.Uri, TargetMode.Internal, Rel + "slideLayout", "slideLayout");
                slide.CreateRelationship(note.Uri, TargetMode.Internal, Rel + "notesSlide", "notesSlide");
                note.CreateRelationship(notesMaster.Uri, TargetMode.Internal, Rel + "notesMaster", "notesMaster");
                note.CreateRelationship(slide.Uri, TargetMode.Internal, Rel + "slide", "slide");
                link?.Invoke(slide, master);

                PackagePart Part(string path, XElement xml, string type)
                {
                    PackagePart part = package.CreatePart(new Uri(path, UriKind.Relative), "application/vnd.openxmlformats-officedocument.presentationml." + type + "+xml");
                    using Stream stream = part.GetStream();
                    xml.Save(stream, SaveOptions.DisableFormatting);
                    return part;
                }
            }

            bytes.Position = 0;
            return EmptySlide.SlidePackage.Read(bytes, expectedId, compareStock: true);
        }

        private static void UnknownMarkupTests(EmptySlide.SlidePackage blank, SlideLike live)
        {
            XNamespace future = "urn:future-office-format";
            // Insert foreign content at every XML position, including the recognized Office
            // creation-ID extension. Future markup must survive normalization and force retention.
            foreach (bool notesPage in new[] { false, true })
            {
                XElement page = notesPage ? blank.Notes! : blank.Slide;
                int elements = page.DescendantsAndSelf().Count();
                for (int index = 0; index < elements; index++)
                {
                    foreach (bool attribute in new[] { false, true })
                    {
                        void Insert(XElement root)
                        {
                            XElement at = root.DescendantsAndSelf().ElementAt(index);
                            if (attribute)
                            {
                                at.SetAttributeValue(future + "value", "keep");
                            }
                            else
                            {
                                at.Add(new XElement(future + "content", "keep"));
                            }
                        }

                        EmptySlide.SlidePackage edited = notesPage ? PackageWith(onNotes: Insert) : PackageWith(onSlide: Insert);
                        Report(Invariant($"empty slide package: foreign {(attribute ? "attribute" : "element")} at position {index} on the {(notesPage ? "notes page" : "slide")}"),
                            !edited.Matches(blank, live, allowMoved: true), "");
                    }
                }
            }
        }

        private static void PictureTextTests(SlideLike live)
        {
            XNamespace p = PresentationMl, a = DrawingMl, future = "urn:future-office-format";
            void Picture(XElement page)
            {
                page.Descendants(p + "ph").Single().SetAttributeValue("type", "pic");
                page.Descendants(p + "txBody").Single().Remove();
            }

            XElement pictureText = new(p + "txBody", new XElement(a + "bodyPr"), new XElement(a + "lstStyle"),
                new XElement(a + "p", new XElement(a + "endParaRPr", new XAttribute("lang", "ja-JP"), new XAttribute("altLang", "en-US"))));
            EmptySlide.SlidePackage picture = PackageWith(onSlide: Picture);
            EmptySlide.SlidePackage WithPictureText()
            {
                return PackageWith(onSlide: x =>
                {
                    Picture(x);
                    x.Descendants(p + "sp").Single().Add(new XElement(pictureText));
                });
            }

            Report("empty slide package: PowerPoint's empty proofing paragraph on a picture placeholder can be omitted",
                WithPictureText().Matches(picture, live, allowMoved: false), "");
            XElement end = pictureText.Descendants(a + "endParaRPr").Single();
            end.SetAttributeValue("lang", "fr-FR");
            Report("empty slide package: a proofing language different from the layout keeps the picture placeholder",
                !WithPictureText().Matches(picture, live, allowMoved: true), "");
            end.SetAttributeValue("lang", "ja-JP");
            end.SetAttributeValue(future + "data", "keep");
            Report("empty slide package: even one extra attribute on the optional proofing paragraph keeps the picture placeholder",
                !WithPictureText().Matches(picture, live, allowMoved: true), "");
        }

        private static void SavedFileRecheck(EmptySlide.SlidePackage blank, EmptySlide.SlidePackage changed)
        {
            EmptySlide.Candidate saved = new(new object(), new object(), blank.Fingerprint, blank.Fingerprint);
            EmptySlide.Candidate unsaved = new(new object(), new object(), blank.Fingerprint, savedFingerprint: null);
            Report("empty slide package: an independently changed, newly saved or missing original invalidates an answer",
                saved.StillMatches(blank, blank) && !saved.StillMatches(blank, changed) && !saved.StillMatches(blank, saved: null) &&
                unsaved.StillMatches(blank, saved: null) && !unsaved.StillMatches(blank, blank), "");
        }

        internal sealed class LegacyPresentation(string folder, string file)
        {
            public string Path { get; } = folder;
            public string FullName { get; } = file;
            public int SlideReads { get; private set; }
            public object Slides
            {
                get
                {
                    SlideReads++;
                    throw new InvalidOperationException();
                }
            }
        }

        private static void UnsupportedFormatTest()
        {
            LegacyPresentation legacy = new("documents", "existing.ppt");
            Report("empty slide: legacy formats are retained before accessing any slide COM property",
                EmptySlide.Find(legacy, out bool fresh, out int count, out string reason) is null &&
                !fresh && count == 0 && legacy.SlideReads == 0 && reason == "Inspection would require a format conversion.", "");
        }

        private static void EmptySlidePackageTests()
        {
            XNamespace p = PresentationMl, a = DrawingMl, future = "urn:future-office-format";
            EmptySlide.SlidePackage blank = PackageWith();
            SlideLike live = new([new ShapeLike(type: 14, text: false)]);
            static void Change(XElement page, XName name)
            {
                page.Descendants(PresentationMl + "spPr").Single().Add(new XElement(name));
            }

            Report("empty slide package: a separately written blank and the stock design match",
                blank.Matches(PackageWith(), live, allowMoved: false) && blank.SameStockDesign(PackageWith()) && blank.Fingerprint.SequenceEqual(PackageWith().Fingerprint), "");
            foreach (XName name in new[] { a + "blipFill", a + "effectLst", a + "gradFill", a + "noFill", a + "ln", a + "scene3d", future + "content" })
            {
                Report("empty slide package: a local " + name.LocalName + " on either page keeps the slide",
                    !PackageWith(onSlide: page => Change(page, name)).Matches(blank, live, allowMoved: true) &&
                    !PackageWith(onNotes: page => Change(page, name)).Matches(blank, live, allowMoved: true), "");
            }

            UnsupportedFormatTest();
            PictureTextTests(live);
            UnknownMarkupTests(blank, live);

            Report("empty slide package: a new creation ID matches, but an additional attribute on that same element is data",
                PackageWith(onSlide: static page => page.Descendants(Office2010 + "creationId").Single().SetAttributeValue("val", "247")).Matches(blank, live, allowMoved: false) &&
                !PackageWith(onSlide: page => page.Descendants(Office2010 + "creationId").Single().SetAttributeValue(future + "mark", "1")).Matches(blank, live, allowMoved: true), "");
            Report("empty slide package: whitespace text and a renamed placeholder are edits",
                !PackageWith(onSlide: page => page.Descendants(a + "p").Single().Add(new XElement(a + "t", " "))).Matches(blank, live, allowMoved: true) &&
                !PackageWith(onSlide: page => page.Descendants(p + "cNvPr").Single().SetAttributeValue("name", "for review")).Matches(blank, live, allowMoved: true), "");
            Report("empty slide package: a slide name requires a question and a missing notes frame keeps the slide",
                PackageWith(onSlide: page => page.Element(p + "cSld").SetAttributeValue("name", "handout")).Matches(blank, live, allowMoved: true) &&
                !PackageWith(onNotes: page => page.Descendants(p + "sp").Single().Remove()).Matches(blank, live, allowMoved: true), "");
            EmptySlide.SlidePackage customMaster = PackageWith(onMaster: xml => xml.Add(new XElement(p + "cSld", "a shared heading")));
            SavedFileRecheck(blank, customMaster);
            Report("empty slide package: a changed master is shared, distinct from the stock design, and invalidates an earlier answer",
                customMaster.Matches(blank, live, allowMoved: true) && !customMaster.SameStockDesign(blank) &&
                !customMaster.Fingerprint.SequenceEqual(blank.Fingerprint), "");
            Report("empty slide package: an additional print or show setting keeps the slide",
                !PackageWith(onOptions: xml => xml.Add(new XElement(p + "prnPr"))).Matches(blank, live, allowMoved: true), "");

            static bool Rejected(Func<EmptySlide.SlidePackage> inspect)
            {
                try
                {
                    inspect();
                    return false;
                }
                catch (InvalidDataException)
                {
                    return true;
                }
            }

            Report("empty slide package: another slide ID, a custom show and unknown presentation elements are rejected",
                Rejected(static () => PackageWith(expectedId: 260)) && Rejected(() => PackageWith(onDeck: x => x.Add(new XElement(p + "custShowLst")))) &&
                Rejected(() => PackageWith(onDeck: x => x.Add(new XElement(future + "show")))) &&
                Rejected(() => PackageWith(onDeck: x => x.SetAttributeValue(future + "slideRef", 259))), "");
            Report("empty slide package: unknown and incoming relationships and duplicate layout links are rejected",
                Rejected(static () => PackageWith(link: static (slide, master) => slide.CreateRelationship(master.Uri, TargetMode.Internal, "urn:future:data"))) &&
                Rejected(static () => PackageWith(link: static (slide, master) => master.CreateRelationship(slide.Uri, TargetMode.Internal, Rel + "slide"))) &&
                Rejected(static () => PackageWith(link: static (slide, _) => slide.CreateRelationship(slide.GetRelationshipsByType(Rel + "slideLayout").Single().TargetUri, TargetMode.Internal, Rel + "slideLayout"))), "");
            XElement placement = new(a + "xfrm", new XElement(a + "off", new XAttribute("x", 20), new XAttribute("y", 30)),
                new XElement(a + "ext", new XAttribute("cx", 300), new XAttribute("cy", 100)));
            EmptySlide.SlidePackage moved = PackageWith(onSlide: x => x.Descendants(p + "spPr").Single().Add(new XElement(placement)));
            SlideLike drawn = new([new ShapeLike(type: 14, text: false) { Framed = (true, 255) }]);
            Report("empty slide package: pure placement is allowed only on an invisible empty slide placeholder and requires confirmation",
                moved.Matches(blank, live, allowMoved: true) && !moved.Matches(blank, live, allowMoved: false) && !moved.Matches(blank, drawn, allowMoved: true) &&
                !PackageWith(onNotes: x => x.Descendants(p + "spPr").Single().Add(placement)).Matches(blank, live, allowMoved: true), "");
            placement.SetAttributeValue(future + "custom", "retain");
            Report("empty slide package: an unknown transform attribute cannot be discarded as mere placement",
                !PackageWith(onSlide: x => x.Descendants(p + "spPr").Single().Add(placement)).Matches(blank, live, allowMoved: true), "");
            XElement section = new(p + "extLst", new XElement(p + "ext", new XAttribute("uri", "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"),
                new XElement(Office2010 + "sectionLst", new XElement(Office2010 + "section", new XAttribute("name", "Introduction"),
                    new XAttribute("id", "{3C0D327D-8549-497A-93B0-B71A53EEDD27}"),
                    new XElement(Office2010 + "sldIdLst", new XElement(Office2010 + "sldId", new XAttribute("id", 259)))))));
            EmptySlide.SlidePackage sectioned = PackageWith(onDeck: x => x.Add(new XElement(section)));
            XElement guideList = new(p + "ext", new XAttribute("uri", "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"),
                new XElement(XName.Get("sldGuideLst", "http://schemas.microsoft.com/office/powerpoint/2012/main")));
            section.Add(guideList);
            Report("empty slide package: the empty guide list PowerPoint adds on opening a file can coexist with a retained section",
                PackageWith(onDeck: x => x.Add(new XElement(section))).Matches(blank, live, allowMoved: true), "");
            section.Add(new XElement(future + "more"));
            Report("empty slide package: a recognized retained section requires confirmation, and any extension beside it keeps the slide",
                sectioned.Matches(blank, live, allowMoved: true) && !sectioned.SameStockDesign(blank) && Rejected(() => PackageWith(onDeck: x => x.Add(section))), "");
            foreach (string setting in new[] { "sldMasterId", "sldSz" })
            {
                Report("empty slide package: an unfamiliar slide reference inside the " + setting + " is retained",
                    Rejected(() => PackageWith(onDeck: x => x.Descendants(p + setting).Single().Add(new XElement(future + "slideRef", 259)))), "");
            }

            Report("empty slide package: a foreign reference nested inside default paragraph formatting is retained",
                Rejected(() => PackageWith(onDeck: x => x.Add(new XElement(p + "defaultTextStyle", new XElement(a + "defPPr",
                    new XElement(a + "defRPr", new XAttribute(future + "slideRef", 259))))))), "");
            Report("empty slide package: extra information on the slide-list entry is retained",
                Rejected(() => PackageWith(onDeck: x => x.Descendants(p + "sldId").Single().Add(new XElement(future + "data")))), "");
            DeckLike unavailable = new(SavedTo, live);
            Report("empty slide: inability to obtain a snapshot at the production entry point always keeps it",
                EmptySlide.Find(unavailable, out bool fresh, out int count, out string reason) is null && !fresh && count == 0 && reason.Length > 0, "");
        }
    }
}
