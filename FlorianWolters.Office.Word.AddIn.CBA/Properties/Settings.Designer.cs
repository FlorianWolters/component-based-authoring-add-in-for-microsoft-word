﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Dieser Code wurde von einem Tool generiert.
//     Laufzeitversion:4.0.30319.18051
//
//     Änderungen an dieser Datei können falsches Verhalten verursachen und gehen verloren, wenn
//     der Code erneut generiert wird.
// </auto-generated>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "11.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_Template")]
        public string WordTemplateFilename {
            get {
                return ((string)(this["WordTemplateFilename"]));
            }
            set {
                this["WordTemplateFilename"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("XML")]
        public string XMLDirectoryName {
            get {
                return ((string)(this["XMLDirectoryName"]));
            }
            set {
                this["XMLDirectoryName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Graphics")]
        public string GraphicsDirectoryName {
            get {
                return ((string)(this["GraphicsDirectoryName"]));
            }
            set {
                this["GraphicsDirectoryName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(".dotm;.dotx")]
        public string WordTemplateFileExtensions {
            get {
                return ((string)(this["WordTemplateFileExtensions"]));
            }
            set {
                this["WordTemplateFileExtensions"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_LastDirectoryPath")]
        public string DocPropertyNameForLastDirectoryPath {
            get {
                return ((string)(this["DocPropertyNameForLastDirectoryPath"]));
            }
            set {
                this["DocPropertyNameForLastDirectoryPath"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://github.com/FlorianWolters/component-based-authoring-add-in-for-microsoft-" +
            "word")]
        public string HostingServiceUrl {
            get {
                return ((string)(this["HostingServiceUrl"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Component-Based Authoring Application Level Add-In for Microsoft Word")]
        public string ApplicationName {
            get {
                return ((string)(this["ApplicationName"]));
            }
        }
    }
}
