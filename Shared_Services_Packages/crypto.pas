unit crypto;

interface

  uses
    System.SysUtils, System.Classes, DECCipher;

  //Encryption
  function encPass(plaintext : RawByteString) : string;
  function decPass(cipherhex : string) : string;

  const passKey;
  const passIV;

implementation


function encPass(plaintext : RawByteString) : String;
var
 Ciphertext: TBytes;
 pwchar : PWideChar;
begin
 with TCipher_Rijndael.Create do
   try
     Mode := cmCFB8;
     Init(passKey, 16, passIV, 16);
     SetLength(Ciphertext, Length(Plaintext));
     Encode(Plaintext[1], Ciphertext[0], Length(Plaintext));
     pwchar := WideStrAlloc(length(Ciphertext));
     BinToHex(Ciphertext,pwchar,length(Ciphertext));
     result := WideCharToString(pwchar);
   finally
     Free;
   end;
end;


function decPass(cipherhex : string) : string;
var
 Plaintext: RawByteString;
 pwchar : PWideChar;
 binBytes : TBytes;
 len : integer;
begin
  pwchar := PWideChar(cipherhex);
  len := length(cipherhex) div 2;
  setLength(binBytes,len);
  HexToBin(pwchar,binBytes,len);

 with TCipher_Rijndael.Create do
   try
     Mode := cmCFB8;
     Init(passKey, 16, passIV, 16);
     SetLength(Plaintext, Length(binBytes));
     FillChar(Plaintext[1], Length(Plaintext), 0);
     Decode(binBytes[0], Plaintext[1], Length(binBytes));
     result := Plaintext;
   finally
     Free;
   end;
end;


end.
