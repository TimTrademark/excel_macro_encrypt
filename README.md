# Excel Macro Encrypt Files PoC

This Excel macro encrypts all files in a given folder (default = C:\\temp), and can be used as a PoC during pentest engagements.
It is possible to detect office files that include macro's, so it might be worthwile to reject those files if an organization does not have a need for them.


## Notice!

Use at your own risk. The files can be decrypted with the same secret. If you do not know what you are doing, double check that you are only encrypting dummy files.