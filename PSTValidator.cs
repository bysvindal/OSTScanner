using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace OutlookOSTScanner
{
    /// <summary>
    /// Low-level PST/OST file validator
    /// Performs deep structural validation
    /// </summary>
    public class PSTValidator
    {
        private readonly string _filePath;
        private readonly List<string> _errors;
        private readonly List<string> _warnings;

        public IReadOnlyList<string> Errors => _errors.AsReadOnly();
        public IReadOnlyList<string> Warnings => _warnings.AsReadOnly();

        public PSTValidator(string filePath)
        {
            _filePath = filePath;
            _errors = new List<string>();
            _warnings = new List<string>();
        }

        /// <summary>
        /// Perform comprehensive validation of PST/OST file
        /// </summary>
        public bool Validate()
        {
            _errors.Clear();
            _warnings.Clear();

            try
            {
                if (!File.Exists(_filePath))
                {
                    _errors.Add("File does not exist");
                    return false;
                }

                using (FileStream fs = new FileStream(_filePath, FileMode.Open,
                    FileAccess.Read, FileShare.Read))
                {
                    // Step 1: Validate file size
                    if (!ValidateFileSize(fs))
                        return false;

                    // Step 2: Validate and parse header
                    var header = ReadHeader(fs);
                    if (header == null)
                        return false;

                    bool isUnicode = (header.Value.wVer == PSTStructures.PST_VERSION_UNICODE ||
                                     header.Value.wVer == PSTStructures.PST_VERSION_UNICODE_2);

                    // Step 3: Validate header CRC
                    if (!ValidateHeaderCRC(fs, isUnicode))
                    {
                        _errors.Add("Header CRC validation failed");
                        return false;
                    }

                    // Step 4: Validate root structure
                    if (!ValidateRootStructure(fs, isUnicode))
                        return false;

                    // Step 5: Validate B-Trees (NBT/BBT)
                    if (!ValidateBTrees(fs, isUnicode))
                    {
                        _warnings.Add("B-Tree validation issues detected");
                    }

                    // Step 6: Validate allocation maps
                    if (!ValidateAllocationMaps(fs, isUnicode))
                    {
                        _warnings.Add("Allocation map validation issues detected");
                    }
                }

                return _errors.Count == 0;
            }
            catch (Exception ex)
            {
                _errors.Add($"Validation exception: {ex.Message}");
                return false;
            }
        }

        private bool ValidateFileSize(FileStream fs)
        {
            if (fs.Length < PSTStructures.HEADER_SIZE_ANSI)
            {
                _errors.Add($"File too small: {fs.Length} bytes (minimum {PSTStructures.HEADER_SIZE_ANSI} required)");
                return false;
            }

            if (fs.Length == 0)
            {
                _errors.Add("File is empty");
                return false;
            }

            return true;
        }

        private PSTStructures.HEADER? ReadHeader(FileStream fs)
        {
            fs.Seek(0, SeekOrigin.Begin);
            byte[] headerBytes = new byte[Marshal.SizeOf<PSTStructures.HEADER>()];

            if (fs.Read(headerBytes, 0, headerBytes.Length) != headerBytes.Length)
            {
                _errors.Add("Failed to read header");
                return null;
            }

            PSTStructures.HEADER header = BytesToStruct<PSTStructures.HEADER>(headerBytes);

            // Validate magic signature
            if (header.dwMagic != PSTStructures.PST_MAGIC)
            {
                _errors.Add($"Invalid magic signature: 0x{header.dwMagic:X8} (expected 0x{PSTStructures.PST_MAGIC:X8})");
                return null;
            }

            // Validate version
            if (header.wVer != PSTStructures.PST_VERSION_ANSI &&
                header.wVer != PSTStructures.PST_VERSION_ANSI_2 &&
                header.wVer != PSTStructures.PST_VERSION_UNICODE &&
                header.wVer != PSTStructures.PST_VERSION_UNICODE_2)
            {
                _errors.Add($"Unsupported PST version: {header.wVer}");
                return null;
            }

            // Validate client magic
            if (header.wMagicClient != 0x534D) // "SM"
            {
                _warnings.Add($"Invalid client magic: 0x{header.wMagicClient:X4}");
            }

            return header;
        }

        private bool ValidateHeaderCRC(FileStream fs, bool isUnicode)
        {
            fs.Seek(0, SeekOrigin.Begin);
            byte[] headerBytes = new byte[PSTStructures.HEADER_SIZE_ANSI];
            fs.Read(headerBytes, 0, headerBytes.Length);

            // Read stored CRC from offset 0x0004
            uint storedCRC = BitConverter.ToUInt32(headerBytes, 0x0004);

            // Calculate CRC over 471 bytes starting at offset 0x0008
            uint calculatedCRC = CalculateCRC32(headerBytes, 0x0008, 471);

            if (storedCRC != calculatedCRC)
            {
                _errors.Add($"Header CRC mismatch: stored=0x{storedCRC:X8}, calculated=0x{calculatedCRC:X8}");
                return false;
            }

            return true;
        }

        private bool ValidateRootStructure(FileStream fs, bool isUnicode)
        {
            fs.Seek(0x00A0, SeekOrigin.Begin);

            if (isUnicode)
            {
                byte[] rootBytes = new byte[Marshal.SizeOf<PSTStructures.ROOT_UNICODE>()];
                if (fs.Read(rootBytes, 0, rootBytes.Length) != rootBytes.Length)
                {
                    _errors.Add("Failed to read Unicode ROOT structure");
                    return false;
                }

                var root = BytesToStruct<PSTStructures.ROOT_UNICODE>(rootBytes);

                // Validate file EOF matches actual file size
                if ((long)root.ibFileEof != fs.Length)
                {
                    _warnings.Add($"File EOF mismatch: header={root.ibFileEof}, actual={fs.Length}");
                }

                return true;
            }
            else
            {
                byte[] rootBytes = new byte[Marshal.SizeOf<PSTStructures.ROOT_ANSI>()];
                if (fs.Read(rootBytes, 0, rootBytes.Length) != rootBytes.Length)
                {
                    _errors.Add("Failed to read ANSI ROOT structure");
                    return false;
                }

                var root = BytesToStruct<PSTStructures.ROOT_ANSI>(rootBytes);

                // Validate file EOF
                if (root.ibFileEof != fs.Length)
                {
                    _warnings.Add($"File EOF mismatch: header={root.ibFileEof}, actual={fs.Length}");
                }

                return true;
            }
        }

        private bool ValidateBTrees(FileStream fs, bool isUnicode)
        {
            // This would involve walking the NBT and BBT structures
            // Complex implementation - checking page signatures, CRCs, and tree integrity
            // Simplified for this example
            _warnings.Add("B-Tree validation not fully implemented (simplified check)");
            return true;
        }

        private bool ValidateAllocationMaps(FileStream fs, bool isUnicode)
        {
            // Validate AMap and PMap structures
            // Simplified for this example
            return true;
        }

        /// <summary>
        /// Calculate CRC32 checksum (PST uses specific CRC32 algorithm)
        /// </summary>
        private uint CalculateCRC32(byte[] data, int offset, int length)
        {
            uint crc = 0;

            // PST uses a specific CRC32 table
            uint[] crcTable = GenerateCRC32Table();

            for (int i = offset; i < offset + length; i++)
            {
                byte index = (byte)((crc & 0xFF) ^ data[i]);
                crc = (crc >> 8) ^ crcTable[index];
            }

            return crc;
        }

        private uint[] GenerateCRC32Table()
        {
            uint[] table = new uint[256];
            uint polynomial = 0xEDB88320;

            for (uint i = 0; i < 256; i++)
            {
                uint crc = i;
                for (int j = 8; j > 0; j--)
                {
                    if ((crc & 1) == 1)
                        crc = (crc >> 1) ^ polynomial;
                    else
                        crc >>= 1;
                }
                table[i] = crc;
            }

            return table;
        }

        private T BytesToStruct<T>(byte[] bytes) where T : struct
        {
            GCHandle handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                return Marshal.PtrToStructure<T>(handle.AddrOfPinnedObject());
            }
            finally
            {
                handle.Free();
            }
        }
    }
}