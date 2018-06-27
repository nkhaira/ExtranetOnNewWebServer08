#region VSS Data
/*
 * VSS Data
 * ----------------------------------------------------------------------------
 * $Source: \\dtmevtvsdv15\D:\Extranet\FileHandling
 * $Author: DANAHERTM\pdeshpan $
 * $Revision: 1.0 $
 * $Date: 2007/08/06 09:41:09 $
 * $Log: FileDataBytes.cs
*/
#endregion
using System;
using System.Collections.Generic;
using System.Text;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using System.IO;
[assembly: ApplicationName("FileHandling")]
[assembly: ApplicationActivation(ActivationOption.Server)]
[assembly: ApplicationAccessControl(false,AccessChecksLevel = AccessChecksLevelOption.ApplicationComponent)]
/// <summary>
/// This class is used to read file bytes in chunks,eg - 10000 bytes at a time.
/// This component gets called from asp page.
/// </summary>
namespace FileHandling
{
    public class FileDataBytes
    {
        #region Private Variables
            private Stream _objFileStream = null;
            private byte[] _abytBuffer;
            private string _strFileName;
            private int    _intChunk;
            private Int32  _intFileSize;
            private Int32  _intByteCount=0;
            private Int32  _intOffSet = 0;
        # endregion
        #region Properties
        /// <summary>
        /// Path of the file from which bytes are read.
        /// </summary>
        public  string strFileName
        {
            get { return _strFileName; }
            set {
                    _strFileName = value;
                    _objFileStream = new FileStream(_strFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                    _intFileSize = Convert.ToInt32(_objFileStream.Length);
                    _abytBuffer = new byte[_intChunk];
                }
        }
        /// <summary>
        /// No of bytes to read at one single time.
        /// </summary>
        public int intChunk
        {
            get { return _intChunk; }
            set { _intChunk = value; }
        }
        /// <summary>
        /// Size of the File in bytes.
        /// </summary>
        public Int32 intFileSize
        {
            get { return _intFileSize; }
            set { _intFileSize = value; }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Function that reads chunk of data in bytes from a file and returns it to the
        /// asp page.
        /// </summary>
        /// <returns>Byte Array containing number of bytes read from a file.</returns>
        public byte[] ReadBytes()
        {
            bool blnClose = false;
            try
            {
                if (intFileSize > intChunk)
                {
                    _abytBuffer = new Byte[_abytBuffer.Length];
                }
                else
                {
                    _abytBuffer = new byte[intFileSize];
                    blnClose = true;
                }
                _intByteCount = _objFileStream.Read(_abytBuffer, _intOffSet, _abytBuffer.Length);
                intFileSize = intFileSize - intChunk;
                if (blnClose==true)
                {
                    _objFileStream.Close();
                    _objFileStream.Dispose();
                }
                return (_abytBuffer);
            }
            catch (Exception exReadData)
            {
                _objFileStream.Close();
                _objFileStream.Dispose();
                return null;
            }
        }
        #endregion
    }
}

