{***********UNITE*************************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 30/11/2001
Modifié le ... : 03/12/2001
Description .. : Objet TFileTools pour gestion de fichiers :
Suite ........ : - présence du lecteur
Suite ........ : - espace libre sur un disque
Suite ........ : - espace total (taille du disque)
Suite ........ : - nom de volume
Suite ........ : (Fonctionnalités extraites de la toz)
Mots clefs ... :
*****************************************************************}
unit galFileTools;

interface

uses Windows,Classes,SysUtils,Dialogs,Forms,Controls,FileCtrl,ShellAPI,Messages
     ,uTob,hent1,hmsgbox,hctrls ;

type
  TTypeSupport = (tsUnknow, tsRemovable, tsFixe, tsNetwork, tsCdrom) ;

Type  LargeInt =  Int64;
Type pLargeInt = ^Int64;

Type
  TFileTools = class (TObject)
  private
        FFileName            : string ; // Nom d'un fichier repère
        FFileDir             : string ; // Répertoire du fichier
        FDrive               : string ; // Lecteur du fichier
        FSizeOfDisk          : LargeInt ; // Taille du disque FDrive
        FFreeOnDisk          : LargeInt ;
        FVolumeName          : string ;  // Nom du volume FDrive
        FDiskSerial          : DWord; // N° de série du disque
        procedure DiskFreeAndSize();
  protected
  public
        function LoadFileProp (Fname : string) : boolean ;
        function GetTypeSupport : TTypeSupport ;
        function IsDiskPresent: Boolean;
        property FileName     : string         read FFileName     ;
        property FileDir      : string         read FFileDir      ;
        property Drive        : string         read FDrive        ;
        property SizeOfDisk   : LargeInt       read FSizeOfDisk   ;
        property VolumeName   : string         read FVolumeName   ;
        property DiskSerial   : DWord          read FDiskSerial   ;
        property FreeOnDisk   : LargeInt       read FFreeOnDisk   ;

        constructor Create ;
        destructor Destroy ; override ;
  end ;


/////////////// IMPLEMENTATION ///////////////
implementation


constructor TFileTools.Create ;
begin
  inherited Create ;
  FFilename         := '';
  FFileDir          := '';
end ;


// détermine et charge les caractéristiques du disque sur lequel se trouve le fichier
// accepte sans pb un chemin de répertoire au lieu d'un fichier
function TFileTools.LoadFileProp (Fname : string ) : boolean ;
begin
  {* Valeur de retour par défaut *}
  Result := False ;

  {* On extrait si présent le répertoire et le drive *}
  FDrive := ExtractFileDrive (Fname) ;

  {* On extrait le répertoire du fichier *}
  FFileDir := ExtractFilePath (Fname) ;

  if GetTypeSupport = tsCdrom then
    begin PGIInfo('Le lecteur '+FDrive+' est un CD-Rom !', TitreHalley); exit; end;

  if Not IsDiskPresent then
    begin PGIInfo('Le lecteur n''est pas présent.', TitreHalley); exit; end;

  {* On extrait le nom du fichier *}
  FFileName := ExtractFileName (Fname) ;

end ;


function TFileTools.GetTypeSupport: TTypeSupport;
begin
  Result := tsUnknow ;
  case GetDriveType (pchar (FDrive)) of
       0,1,DRIVE_RAMDISK : result := tsUnknow ;
       DRIVE_REMOVABLE : result := tsRemovable ;
       DRIVE_FIXED     : result := tsFixe ;
       DRIVE_REMOTE    : result := tsNetwork ;
       DRIVE_CDROM     : result := tsCdrom ;
       end ;
end;

{*
Cette procédure renseigne FSizeOfDisk, FFreeOnDisk.
*}
procedure TFileTools.DiskFreeAndSize();   // RCV150199
var
   GetDiskFreeSpaceEx: function( RootName: pChar; FreeForCaller, TotNoOfBytes: pLargeInt; TotNoOfFreeBytes: pLargeInt ): BOOL; stdcall;
   SectorsPCluster, BytesPSector, FreeClusters, TotalClusters: DWORD;
   LDiskFree, LSizeOfDisk: LargeInt;
   Lib: THandle;
begin
   SectorsPCluster := 0 ;
   BytesPSector := 0 ;
   FreeClusters := 0 ;
   TotalClusters := 0 ;
   LDiskFree   := -1;
   LSizeOfDisk := -1;
   Lib := GetModuleHandle( 'Kernel32' );
   If Lib <> 0 then
   begin
      @GetDiskFreeSpaceEx := GetProcAddress( Lib, 'GetDiskFreeSpaceExA' );
      if ( @GetDiskFreeSpaceEx <> nil ) then   // We probably have W95+OSR2 or better.
      begin
         if Not GetDiskFreeSpaceEx( pChar( FDrive ), @LDiskFree, @LSizeOfDisk, nil ) then
         begin
            LDiskFree   := -1;
            LSizeOfDisk := -1;
         end;
         FreeLibrary( Lib );
         @GetDiskFreeSpaceEx := nil;
      end
   end;
   if ( LDiskFree = -1 ) then   // We have W95 original or W95+OSR1 or an error.
   begin   // We use this because DiskFree/Size don't support UNC drive names.
      if GetDiskFreeSpace( pChar( FDrive ), SectorsPCluster, BytesPSector, FreeClusters, TotalClusters )  then
      begin
         LDiskFree   := LargeInt( BytesPSector )
                      * SectorsPCluster * FreeClusters ;
                      // * LargeInt( SectorsPCluster ) * LargeInt( FreeClusters );
         LSizeOfDisk := LargeInt( BytesPSector )
                      * SectorsPCluster * TotalClusters ;
      end;
   end;

   FFreeOnDisk := LDiskFree;
   FSizeOfDisk := LSizeOfDisk;
end;


{*
Cette fct permet de connaître la validité du disque FDrive
Renseigne FVolumeName avec le label du disque
Renseigne FSizeOfDisk avec la taille du disque => par l'intermé. de DiskFreeAndSize
Renseigne FFreeOnDisk avec l'espace libre  => par l'intermé. de DiskFreeAndSize
Sinon retourne False
*}
function TFileTools.IsDiskPresent: Boolean;
var
   SysFlags, OldErrMode: DWord;
   NamLen:               Cardinal;
   SysLen:               DWord;
   Buf:                  String;
   VolNameAry:           Array[0..255] of Char;
   Num:                  Integer;
   Bits:                 Set of 0..25;
   DriveLetter:          Char;
begin
   NamLen      := 255;
   SysLen      := 255;
   FSizeOfDisk := 0;
   FVolumeName := '';
   Result      := False;
   DriveLetter := UpperCase( FDrive )[1];
   Buf         := DriveLetter + ':\';

// Modif CA - 30/05/2002 - Pb si lancement sur serveur : on a le chemin
// de la manière suivante \\Serveur ... ==> taille mal calculée ==> mise à jour impossible.
// Fait à l'arrache ... Mais bon !
   if (DriveLetter = '\') then
      begin
//      Result := True ;
//      Exit ;
//      end ;
      end
else //  Ajout CA - 30/05/2002
  begin //  Ajout CA - 30/05/2002
   if (DriveLetter < 'A') or (DriveLetter > 'Z') then
     begin PGIInfo(DriveLetter + ' n''est pas un lecteur !', TitreHalley); exit; end;

  {* Est-ce que FDrive est OK ? *}
   Integer(Bits) := GetLogicalDrives();
   Num := Ord( DriveLetter ) - Ord( 'A' );
   if NOT (Num in Bits) then
     begin PGIInfo('Le lecteur '+DriveLetter+' n''est pas connecté.', TitreHalley); exit; end;
  end; //  Ajout CA - 30/05/2002

   {* Désactivation temporaire des erreurs critiques *}
   OldErrMode := SetErrorMode( SEM_FAILCRITICALERRORS );

   {* Depuis la version v1.52c, il n'y a plus d'exception raisée *}
   if NOT GetVolumeInformation( pChar( Buf ), VolNameAry, NamLen, @FDiskSerial, SysLen, SysFlags, nil, 0 ) then
     begin
     // W'll get this if there is a disk but it is not or wrong formatted
     // so this disk can only be used when we also want formatting.
     if (GetLastError() = 31) then
        begin PGIInfo('Le lecteur '+DriveLetter+' n''est pas formatté.', TitreHalley); exit; end;
     end;

   {* On récupère la nom du volume du disque en question *}
   FVolumeName := VolNameAry;

   {* Récupération de l'espace disque libre et de l'espace disque total *}
   DiskFreeAndSize();

   {* On réactive le mode erreur critique *}
   SetErrorMode( OldErrMode );

   {* Faire attention à des disques réseaux au format UNC. Dans certains, notament sous Windows 95,
      une valeur -1 est renvoyée.
   *}
   if (DriveLetter = '\') or ( (DriveLetter <> '\') and (FSizeOfDisk <> -1) ) then
      Result := True;
end;


destructor TFileTools.Destroy;
begin
  inherited;

end;

////////////////////////////////////////////////////////////////////////////

end.



