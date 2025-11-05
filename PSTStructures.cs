using System;
using System.Runtime.InteropServices;

namespace OutlookOSTScanner
{
    /// <summary>
    /// PST/OST file format structures based on MS-PST specification
    /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst
    /// </summary>
    public static class PSTStructures
    {
        // Magic numbers
        public const uint PST_MAGIC = 0x4E444221; // "!BDN" in little-endian
        public const ushort PST_VERSION_ANSI = 14;
        public const ushort PST_VERSION_ANSI_2 = 15;
        public const ushort PST_VERSION_UNICODE = 23;
        public const ushort PST_VERSION_UNICODE_2 = 36; // Outlook 2013+

        // File offsets
        public const int HEADER_SIZE_ANSI = 564;
        public const int HEADER_SIZE_UNICODE = 564;

        /// <summary>
        /// HEADER structure for PST/OST files
        /// Offset 0x0000, size varies by version
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct HEADER
        {
            public uint dwMagic;           // 0x0000: Must be "!BDN" (0x4E444221)
            public uint dwCRCPartial;      // 0x0004: CRC of 471 bytes starting at 0x0008
            public ushort wMagicClient;    // 0x0008: Must be "SM" (0x534D)
            public ushort wVer;            // 0x000A: File format version
            public ushort wVerClient;      // 0x000C: Client version
            public byte bPlatformCreate;   // 0x000E: Platform (0x01 = Win)
            public byte bPlatformAccess;   // 0x000F: Platform access
            public uint dwReserved1;       // 0x0010
            public uint dwReserved2;       // 0x0014
        }

        /// <summary>
        /// ROOT structure for ANSI PST (version 14/15)
        /// Offset 0x00A0
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct ROOT_ANSI
        {
            public uint dwReserved;        // 0x00A0
            public uint ibFileEof;         // 0x00A4: File EOF offset
            public uint ibAMapLast;        // 0x00A8: AMap last
            public uint cbAMapFree;        // 0x00AC: Free space in AMap
            public uint cbPMapFree;        // 0x00B0: Free space in PMap
            // NBT and BBT page structures follow
            public BBTENTRY bbtRoot;       // 0x00B4: Root BBT entry
            public NBTENTRY nbtRoot;       // 0x00BC: Root NBT entry
        }

        /// <summary>
        /// ROOT structure for Unicode PST (version 23/36)
        /// Offset 0x00A0
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct ROOT_UNICODE
        {
            public uint dwReserved;        // 0x00A0
            public ulong ibFileEof;        // 0x00A4: File EOF offset (64-bit)
            public ulong ibAMapLast;       // 0x00AC: AMap last (64-bit)
            public ulong cbAMapFree;       // 0x00B4: Free space in AMap
            public ulong cbPMapFree;       // 0x00BC: Free space in PMap
            // NBT and BBT structures
            public BBTENTRY_UNICODE bbtRoot;  // Root BBT entry
            public NBTENTRY_UNICODE nbtRoot;  // Root NBT entry
        }

        /// <summary>
        /// BBT Entry (Block BTree) - ANSI
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BBTENTRY
        {
            public ulong bref;             // Block reference
            public ushort cb;              // Count of bytes
            public ushort cRef;            // Reference count
        }

        /// <summary>
        /// BBT Entry (Block BTree) - Unicode
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BBTENTRY_UNICODE
        {
            public ulong bref;             // Block reference (64-bit)
            public ushort cb;              // Count of bytes
            public ushort cRef;            // Reference count
            public uint dwPadding;         // Padding for alignment
        }

        /// <summary>
        /// NBT Entry (Node BTree) - ANSI
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct NBTENTRY
        {
            public uint nid;               // Node ID
            public ulong bidData;          // Data block ID
            public ulong bidSub;           // Sub-node block ID
            public uint nidParent;         // Parent node ID
        }

        /// <summary>
        /// NBT Entry (Node BTree) - Unicode
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct NBTENTRY_UNICODE
        {
            public ulong nid;              // Node ID (64-bit)
            public ulong bidData;          // Data block ID
            public ulong bidSub;           // Sub-node block ID
            public uint dwPadding;         // Padding
        }

        /// <summary>
        /// Page trailer for validation
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct PAGETRAILER
        {
            public byte ptype;             // Page type
            public byte ptypeRepeat;       // Page type repeat
            public ushort wSig;            // Signature
            public uint dwCRC;             // CRC of page
            public ulong bid;              // Block ID
        }

        // Page types
        public const byte ptypeBBT = 0x80;    // BBT page
        public const byte ptypeNBT = 0x81;    // NBT page
        public const byte ptypeFMap = 0x82;   // Free Map page
        public const byte ptypePMap = 0x83;   // Page Map page
        public const byte ptypeAMap = 0x84;   // Allocation Map page
        public const byte ptypeFPMap = 0x85;  // Free Page Map page
        public const byte ptypeDL = 0x86;     // Density List page
    }
}