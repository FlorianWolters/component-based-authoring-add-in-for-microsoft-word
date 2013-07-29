//------------------------------------------------------------------------------
// <copyright file="Settings.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Properties
{
    using System.ComponentModel;
    using System.Configuration;

    /// <summary>
    /// The class <see cref="Settings"/> allows to handle the events of the <see
    /// cref="ApplicationSettingsBase"/> class.
    /// </summary>
    internal sealed partial class Settings
    {
        /// <summary>
        ///  Initializes a new instance of the <see cref="Settings"/> class.
        /// </summary>
        public Settings()
        {
            ////this.SettingChanging += this.SettingChangingEventHandler;
            ////this.PropertyChanged += this.PropertyChangedEventHandler;
            ////this.SettingsLoaded += this.SettingsLoadedEventHandler;
            ////this.SettingsSaving += this.SettingsSavingEventHandler;
        }
        
        ////private void SettingChangingEventHandler(object sender, SettingChangingEventArgs e)
        ////{
        ////}

        ////private void PropertyChangedEventHandler(object sender, PropertyChangedEventArgs e)
        ////{
        ////}

        ////private void SettingsLoadedEventHandler(object sender, SettingsLoadedEventArgs e)
        ////{
        ////}

        ////private void SettingsSavingEventHandler(object sender, CancelEventArgs e)
        ////{
        ////}
    }
}
