﻿// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using NuGet.ProjectManagement;

namespace NuGet.PackageManagement.VisualStudio
{
    /// <summary>
    /// Class that manages the binding redirect config section
    /// </summary>
    public class BindingRedirectManager
    {
        private static readonly XName AssemblyBindingName = AssemblyBinding.GetQualifiedName("assemblyBinding");
        private static readonly XName DependentAssemblyName = AssemblyBinding.GetQualifiedName("dependentAssembly");

        private string ConfigurationFile { get; set; }
        private IMSBuildNuGetProjectSystem MSBuildNuGetProjectSystem { get; set; }

        public BindingRedirectManager(string configurationFile, IMSBuildNuGetProjectSystem msBuildNuGetProjectSystem)
        {
            if (String.IsNullOrEmpty(configurationFile))
            {
                throw new ArgumentException(ProjectManagement.Strings.Argument_Cannot_Be_Null_Or_Empty, "configurationFile");
            }
            if (msBuildNuGetProjectSystem == null)
            {
                throw new ArgumentNullException("msBuildNuGetProjectSystem");
            }

            ConfigurationFile = configurationFile;
            MSBuildNuGetProjectSystem = msBuildNuGetProjectSystem;
        }

        public void AddBindingRedirects(IEnumerable<AssemblyBinding> bindingRedirects)
        {
            if (bindingRedirects == null)
            {
                throw new ArgumentNullException("bindingRedirects");
            }

            // Do nothing if there are no binding redirects to add, bail out
            if (!bindingRedirects.Any())
            {
                return;
            }

            WriteToXmlFile(
                Path.Combine(MSBuildNuGetProjectSystem.ProjectFullPath, ConfigurationFile),
                document =>
                {
                    // Get the runtime element
                    XElement runtime = document.Root.Element("runtime");

                    if (runtime == null)
                    {
                        // Add the runtime element to the configuration document
                        runtime = new XElement("runtime");
                        document.Root.AddIndented(runtime);
                    }

                    // Get all of the current bindings in config
                    ILookup<AssemblyBinding, XElement> currentBindings = GetAssemblyBindings(document);

                    XElement assemblyBindingElement = null;
                    foreach (var bindingRedirect in bindingRedirects)
                    {
                        // Look to see if we already have this in the list of bindings already in config.
                        if (currentBindings.Contains(bindingRedirect))
                        {
                            var existingBindings = currentBindings[bindingRedirect];
                            if (existingBindings.Any())
                            {
                                // Remove all but the first assembly binding elements
                                foreach (var bindingElement in existingBindings.Skip(1))
                                {
                                    RemoveElement(bindingElement);
                                }

                                UpdateBindingRedirectElement(existingBindings.First(), bindingRedirect);
                                // Since we have a binding element, the assembly binding node (parent node) must exist. We don't need to do anything more here.
                                continue;
                            }
                        }

                        if (assemblyBindingElement == null)
                        {
                            // Get an assembly binding element to use
                            assemblyBindingElement = GetAssemblyBindingElement(runtime);
                        }
                        // Add the binding to that element

                        assemblyBindingElement.AddIndented(bindingRedirect.ToXElement());
                    }
                });

            // Ensure the file is added to the project
            MSBuildNuGetProjectSystem.AddExistingFile(ConfigurationFile);
        }

        public void RemoveBindingRedirects(IEnumerable<AssemblyBinding> bindingRedirects)
        {
            if (bindingRedirects == null)
            {
                throw new ArgumentNullException("bindingRedirects");
            }

            // Do nothing if there are no binding redirects to remove, bail out
            if (!bindingRedirects.Any())
            {
                return;
            }

            WriteToXmlFile(
                Path.Combine(MSBuildNuGetProjectSystem.ProjectFullPath, ConfigurationFile),
                document =>
                {
                    // Get all of the current bindings in config
                    ILookup<AssemblyBinding, XElement> currentBindings = GetAssemblyBindings(document);

                    if (!currentBindings.Any())
                    {
                        return;
                    }

                    foreach (var bindingRedirect in bindingRedirects)
                    {
                        if (currentBindings.Contains(bindingRedirect))
                        {
                            foreach (var bindingElement in currentBindings[bindingRedirect])
                            {
                                RemoveElement(bindingElement);
                            }
                        }
                    }
                });
        }

        private static void WriteToXmlFile(string fileName, Action<XDocument> performWrite)
        {
            var serviceProvider = ServiceLocator.GetInstance<IServiceProvider>();
            using (var file = BufferedTextFile.CreateBufferedTextFile(serviceProvider, fileName))
            {
                var buffer = file.GetTextBuffer(ensureWritable: true);
                using (var edit = buffer.CreateEdit())
                {
                    var currentText = buffer.CurrentSnapshot.GetText();
                    var document = XDocument.Parse(currentText);

                    performWrite(document);

                    edit.Replace(0, currentText.Length, GetStringContents(document));
                    edit.Apply();
                }

                file.SaveIfDirty();
            }
        }

        private static string GetStringContents(XDocument document)
        {
            var stringBuilder = new StringBuilder();
            using (var writer = new CustomEncodingStringWriter(stringBuilder, document.Declaration?.Encoding))
            {
                document.Save(writer);
                writer.Flush();
            }

            return stringBuilder.ToString();
        }

        private static void RemoveElement(XElement element)
        {
            // Hold onto the parent element before removing the element
            XElement parentElement = element.Parent;

            // Remove the element from the document if we find a match
            element.RemoveIndented();

            if (!parentElement.HasElements)
            {
                parentElement.RemoveIndented();
            }
        }

        private static XElement GetAssemblyBindingElement(XElement runtime)
        {
            // Pick the first assembly binding element or create one if there aren't any
            XElement assemblyBinding = runtime.Elements(AssemblyBindingName).FirstOrDefault();
            if (assemblyBinding != null)
            {
                return assemblyBinding;
            }

            assemblyBinding = new XElement(AssemblyBindingName);
            runtime.AddIndented(assemblyBinding);

            return assemblyBinding;
        }

        private static ILookup<AssemblyBinding, XElement> GetAssemblyBindings(XDocument document)
        {
            XElement runtime = document.Root.Element("runtime");

            IEnumerable<XElement> assemblyBindingElements = Enumerable.Empty<XElement>();
            if (runtime != null)
            {
                assemblyBindingElements = GetAssemblyBindingElements(runtime);
            }

            // We're going to need to know which element is associated with what binding for removal
            var assemblyElementPairs = from dependentAssemblyElement in assemblyBindingElements
                                       select new
                                       {
                                           Binding = AssemblyBinding.Parse(dependentAssemblyElement),
                                           Element = dependentAssemblyElement
                                       };

            // Return a mapping from binding to element
            return assemblyElementPairs.ToLookup(p => p.Binding, p => p.Element);
        }

        private static IEnumerable<XElement> GetAssemblyBindingElements(XElement runtime)
        {
            return runtime.Elements(AssemblyBindingName)
                .Elements(DependentAssemblyName);
        }

        private static void UpdateBindingRedirectElement(XElement element, AssemblyBinding bindingRedirect)
        {
            var bindingRedirectElement = element.Element(AssemblyBinding.GetQualifiedName("bindingRedirect"));
            // Since we've successfully parsed this node, it has to be valid and this child must exist.
            Debug.Assert(bindingRedirectElement != null);
            bindingRedirectElement.Attribute("oldVersion").SetValue(bindingRedirect.OldVersion);
            bindingRedirectElement.Attribute("newVersion").SetValue(bindingRedirect.NewVersion);
        }

        private class CustomEncodingStringWriter : StringWriter
        {
            private Encoding _encoding;

            public CustomEncodingStringWriter(StringBuilder stringBuilder, string encodingWebName)
                : base(stringBuilder)
            {
                _encoding = Encoding.UTF8;

                if (encodingWebName != null)
                {
                    try
                    {
                        _encoding = Encoding.GetEncoding(encodingWebName);
                    }
                    catch (ArgumentException)
                    {
                        // In case of an invalid encoding name, use UTF-8.
                    }
                }
            }

            public override Encoding Encoding { get { return _encoding; } }
        }
    }
}
