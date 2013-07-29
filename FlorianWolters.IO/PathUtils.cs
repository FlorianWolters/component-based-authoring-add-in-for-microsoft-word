//------------------------------------------------------------------------------
// <copyright file="PathUtils.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.IO
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;

    /// <summary>
    /// The static class <see cref="PathUtils"/> contains utility methods for
    /// file path manipulation.
    /// </summary>
    /// <remarks>
    /// The source code is taken from <a
    /// href="http://stackoverflow.com/questions/275689/how-to-get-relative-path-from-absolute-path">this</a>
    /// StackOverflow question.
    /// </remarks>
    public static class PathUtils
    {
        /// <summary>
        /// The maximum number of characters for a file path.
        /// </summary>
        public const int MaxPath = 260; 

        /// <summary>
        /// Signals that the file is a directory.
        /// </summary>
        private const int FileAttributeDirectory = 0x10;

        /// <summary>
        /// Signals that the file is a normal file.
        /// </summary>
        private const int FileAttributeNormal = 0x80;

        /// <summary>
        /// Retrieves the relative path from one normal file or directory to
        /// another.
        /// </summary>
        /// <param name="fromPath">The start of the relative path.</param>
        /// <param name="toPath">The endpoint of the relative path.</param>
        /// <returns>The relative path from the start to the endpoint.</returns>
        /// <exception cref="ArgumentNullException">If at least one argument is <c>null</c> or an empty string.</exception>
        /// <exception cref="ArgumentException">If the two path do not have a common prefix.</exception>
        /// <exception cref="FileNotFoundException">If at least one path does not exist.</exception>
        public static string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath))
            {
                throw new ArgumentNullException("fromPath");
            }

            if (string.IsNullOrEmpty(toPath))
            {
                throw new ArgumentNullException("toPath");
            }

            int fromAttr = GetPathAttribute(fromPath);
            int toAttr = GetPathAttribute(toPath);

            StringBuilder path = new StringBuilder(MaxPath);

            if (0 == NativeMethods.PathRelativePathTo(
                path,
                fromPath,
                fromAttr,
                toPath,
                toAttr))
            {
                throw new ArgumentException(
                    "The paths must have a common prefix.");
            }

            return path.ToString();
        }

        /// <summary>
        /// Retrieves the file attribute for the specified relative or absolute
        /// file path.
        /// </summary>
        /// <param name="path">The file path to check.</param>
        /// <returns>
        /// <see cref="FileAttributeDirectory"/> if the file path specifies a
        /// directory or <see cref="FileAttributeNormal"/> if the file path
        /// specifies a normal file.
        /// </returns>
        /// <exception cref="FileNotFoundException">If at least one path does not exist.</exception>
        private static int GetPathAttribute(string path)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            if (directoryInfo.Exists)
            {
                return FileAttributeDirectory;
            }

            FileInfo fileInfo = new FileInfo(path);
            if (fileInfo.Exists)
            {
                return FileAttributeNormal;
            }

            throw new FileNotFoundException();
        }

        /// <summary>
        /// The class <see cref="NativeMethods"/> contains native methods, used
        /// by the class <see cref="PathUtils"/>.
        /// </summary>
        internal static class NativeMethods
        {
            /// <summary>
            /// Creates a relative path from one file or folder to another.
            /// </summary>
            /// <param name="pszPath">A pointer to a string that receives the relative path. This buffer must be at least <c>MAX_PATH</c> characters in size.</param>
            /// <param name="pszFrom">A pointer to a null-terminated string of maximum length <c>MAX_PATH</c> that contains the path that defines the start of the relative path.</param>
            /// <param name="dwAttrFrom">The file attributes of <c>pszFrom</c>. If this value contains <c>FILE_ATTRIBUTE_DIRECTORY</c>, <c>pszFrom</c> is assumed to be a directory; otherwise, <c>pszFrom</c> is assumed to be a file.</param>
            /// <param name="pszTo">A pointer to a null-terminated string of maximum length <c>MAX_PATH</c> that contains the path that defines the endpoint of the relative path.</param>
            /// <param name="dwAttrTo">The file attributes of <c>pszTo</c>. If this value contains <c>FILE_ATTRIBUTE_DIRECTORY</c>, <c>pszTo</c> is assumed to be directory; otherwise, <c>pszTo</c> is assumed to be a file.</param>
            /// <returns><c>true</c>if successful, or <c>false</c> otherwise.</returns>
            /// <remarks><a href="http://msdn.microsoft.com/windows/desktop/bb773740.aspx">MSDN</a></remarks>
            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            internal static extern int PathRelativePathTo(
                StringBuilder pszPath,
                string pszFrom,
                int dwAttrFrom,
                string pszTo,
                int dwAttrTo);
        }
    }
}
