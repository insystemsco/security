@echo off
echo off
rem ******************************************************
rem DESCRIPTION: Displays running encryption executibles
rem found by listing the running tasks and then searching
rem them for the encryption executible's name.
rem SYNTAX: encryptiondetect <drive_letter>
rem EXAMPLE: encryptionrunning
rem NOTE: This command should be run with Administrator
rem privileges.  It may also miss certain encryption or
rem give false positives if exe names have been changed.
rem OPERATING SYSTEM VERSIONS: 2000,XP,Vista,2003,2008
rem ******************************************************

rem This batch file should (eventually) be able to identify
rem executibles associated with the following encryption
rem programs:
rem ArchiCrypt Live
rem BestCrypt
rem BitArmor Data Control
rem BitLocker
rem CGFD
rem Check Point Full Disk Encryption
rem Cross Crypt
rem Crypt Archiver
rem CryptoLoop
rem DiskCryptor
rem DISK Protect (BeCrypt)
rem dm-crypt / cryptsetup
rem dn-crypt / LUKS
rem DriveCrypt
rem DriveSentry
rem E4M
rem e-Capsule Private Safe
rem eCryptfs
rem FileVault
rem FinallySecure
rem FREE CompuSec
rem FreeOTFE
rem GBDE
rem GELI
rem Keyparc
rem Loop-AES
rem PGPDisk
rem Private Disk 
rem R-Crypto
rem MacAfee EndPoint Encryption
rem Safeguard Easy (Utimaco)
rem SafeHouse Professional
rem Scramdisk
rem SecuBox
rem SEcude securenotebook
rem SecureDoc
rem Sentry 2020
rem SpyProof!
rem TrueCrypt

echo **********************************************
echo START DATE AND TIME:
date /t
time /t
echo **********************************************
echo.

echo KNOWN ENCRYPTION EXECUTIBLES CURRENTLY RUNNING ON THE SUSPECT COMPUTER:
echo.

tasklist | findstr /I "ArchiCrypt BestCrypt BitArmor BitLocker CGFD checkpoint CrossCrypt CryptArchiver CryptoLoop DiskCryptor DISKProtect BeCrypt dm-crypt n-crypt DriveCrypt DriveSentry E4M e-Capsule eCryptfs FileVault FinallySecure CompuSec FreeOTFE GBDE GELI Keyparc Loop-AES PGPDisk PrivateDisk R-Crypto EndPoint Safeguard SafeHouse Scramdisk SecuBox SEcude securenotebook SecureDoc Sentry SpyProof TrueCrypt"
echo.
echo.

echo **********************************************
echo END DATE AND TIME:
date /t
time /t
echo **********************************************