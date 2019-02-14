@echo off
echo off
rem ******************************************************
rem DESCRIPTION: Displays windows encrypted files and
rem folders on a logical drive and then searches the drive
rem for file or folder names used by the more popular
rem cryptographic tools.
rem SYNTAX: encryptiondetect <drive_letter>
rem EXAMPLE: encryptiondetect C:
rem NOTE: This command should be run with Administrator
rem privileges.  It may also miss certain encryption or
rem give false positives.
rem encryptionrunning.bat should be run before encryptiondetect.bat
rem as encryptiondetect will cause cipher.exe to run
rem OPERATING SYSTEM VERSIONS: 2000,XP,Vista,2003,2008
rem ******************************************************

echo **********************************************
echo INVESTIGATOR NAME: 
echo **********************************************
echo START DATE AND TIME:
date /t
time /t
echo **********************************************
echo.

echo POSSIBLE WINDOWS-ENCRYPTED FILES AND FOLDERS DETECTED ON %1 DRIVE:
echo.

cipher /s:%1:\ | findstr /b ^.E

echo.
echo.

echo **********************************************
echo.
echo FILES OR FOLDERS CONTAINING THE TERM "CRYPT" IN THEIR NAMES FOUND ON %1 DRIVE:
echo.
dir %1:\* /Q /s | find /I "crypt"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.
dir %1:\* /A:H /Q /s | find /I "crypt"
echo.
echo.


echo **********************************************
echo.
echo POSSIBLE EVIDENCE OF TRUECRYPT FOUND ON %1 DRIVE:
echo.
dir %1:\* /s | find /I "truecrypt"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.
dir %1:\* /A:H /Q /s | find /I "truecrypt"
echo.
echo.

echo **********************************************
echo.
echo POSSIBLE EVIDENCE OF BESTCRYPT FOUND ON %1 DRIVE:
echo.
dir %1:\* /s | find /I "bestcrypt"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.
dir %1:\* /A:H /Q /s | find /I "bestcrypt"
echo.
echo.

echo **********************************************
echo.
echo POSSIBLE EVIDENCE OF SAFEHOUSE ENCRYPTION FOUND ON %1 DRIVE:
echo.
dir %1:\* /s | find /I "safehouse"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.
dir %1:\* /A:H /Q /s | find /I "bestcrypt"
echo.
echo.

echo **********************************************
echo.
echo POSSIBLE EVIDENCE OF PGP FOUND ON %1 DRIVE:
echo.
dir %1:\* /s | find /I "pgp"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.
dir %1:\* /A:H /Q /s | find /I "pgp"
echo.
echo.

echo **********************************************
echo.
echo POSSIBLE EVIDENCE OF STEGANOGRAPHIC SOFTWARE FOUND ON %1 DRIVE:
echo.
dir %1:\* /s | find /I "steg"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.
dir %1:\* /A:H /Q /s | find /I "steg"
echo.
echo.


echo **********************************************
echo.
echo POSSIBLE EVIDENCE OF OTHER ENCRYPTION SOFTWARE FOUND ON %1 DRIVE:
echo.
echo.

dir %1:\* /s | findstr /I "ArchiCrypt BestCrypt BitArmor BitLocker CGFD checkpoint CrossCrypt CryptArchiver CryptoLoop DiskCryptor DISKProtect BeCrypt dm-crypt dn-crypt DriveCrypt DriveSentry E4M e-Capsule eCryptfs FileVault FinallySecure CompuSec FreeOTFE GBDE GELI Keyparc Loop-AES PGPDisk PrivateDisk  R-Crypto EndPoint Safeguard SafeHouse Scramdisk SecuBox SEcude securenotebook SecureDoc Sentry SpyProof Steg TrueCrypt"
echo.
echo ***
echo.
echo Hidden files or folders:
echo.

dir %1:\* /A:H /Q /s | findstr /I "ArchiCrypt BestCrypt BitArmor BitLocker CGFD checkpoint CrossCrypt CryptArchiver CryptoLoop DiskCryptor DISKProtect
 BeCrypt dm-crypt dn-crypt DriveCrypt DriveSentry E4M e-Capsule eCryptfs FileVault FinallySecure CompuSec FreeOTFE GBDE GELI Keyparc Loop-AES PGPDisk PrivateDisk  R-Crypto EndPoint Safeguard SafeHouse Scramdisk SecuBox SEcude securenotebook SecureDoc Sentry SpyProof Steg TrueCrypt"

echo.
echo.

echo **********************************************
echo END DATE AND TIME:
date /t
time /t
echo **********************************************