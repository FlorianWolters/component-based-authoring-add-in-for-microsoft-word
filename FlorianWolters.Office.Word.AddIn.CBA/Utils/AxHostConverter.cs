//------------------------------------------------------------------------------
// <copyright file="AxHostConverter.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Utils
{
    using System.Drawing;
    using System.Windows.Forms;
    using stdole;

    public class AxHostConverter : AxHost
    {
        private AxHostConverter() : base(string.Empty)
        {
        }

        public static IPictureDisp ImageToPictureDisp(Image image)
        {
            return (IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
        }

        public static Image PictureDispToImage(IPictureDisp pictureDisp)
        {
            return AxHost.GetPictureFromIPicture(pictureDisp);
        }
    }
}
