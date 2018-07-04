{***********UNITE*************************************************
Auteur  ...... : Laurent Abélard
Créé le ...... : 19/11/2007
Modifié le ... : 23/01/2008
Description .. : Incrémentation du Numéro de Build d'un fichier de
Suite ........ : ressource .res
Suite ........ : La compilation d'un projet en ligne de commande
Suite ........ : n'incrémente pas le Numéro de Build. Seul IDE
Suite ........ : le fait de façon automatique (si l'option est
Suite ........ : cochée) en reconstruisant le fichier de ressource.
Suite ........ : La méthode généralement utilisée consiste à
Suite ........ : recréer un fichier source de ressource .rc à
Suite ........ : l'aide des informations stockées dans le fichier
Suite ........ : .dof mais cette méthode ne permet pas de conserver
Suite ........ : d'autre ressource que le programmeur aurait stocké
Suite ........ : dans le .res.
Suite ........ : J'ai tester l'usage de la fonction UpdateResource()
Suite ........ : mais cela ne fonctionne pas avec des .res.
Suite ........ : J'opte pour manipuler directement le fichier
Suite ........ : binaire qui est bien documenté par Microsoft en :
Suite ........ : ms-help://MS.VSCC.v80/MS.MSDN.v80/MS.WIN32COM.v10.en/winui/winui/windowsuserinterface/resources/introductiontoresources/resourcefileformats.htm
Suite ........ : http://msdn.microsoft.com/library/en-us/winui/WinUI/WindowsUserInterface/Resources/IntroductiontoResources/ResourceReference/ResourceFunctions/LoadImage.asp
Suite ........ : Transformée en http://msdn2.microsoft.com/fr-fr/library/ms648007.aspx
Mots clefs ... : Ressource, .res, .rc
*****************************************************************}
Unit IncBuildSources;

Interface

Type
  BUILD_OU_VERSION = ( BUILD_FORCE=1, VERSION_FORCE );

  Function IncrementerBuild( CheminRes : String; CheminXml : String = ''; ValeurForcee : String = '';
                             BuildOuVersion : BUILD_OU_VERSION = BUILD_FORCE; VersionProduit : String = '' ) : String;

Implementation
Uses
  Windows,      //Pour VS_FIXEDFILEINFO
  SysUtils,     //Pour ChangeFileExt(), DeleteFile(), RenameFile()
  //hctrls,       //Pour ReadTokenPipe() de l'Agl
  Classes;      //Pour TFileStream;

Type
  TDWORD  = Array [0..MaxInt DIV SizeOf(DWORD)-1] Of DWORD; //Tableau illimité de DWORD.
  PTDWORD = ^TDWORD;

  //Déscription de la structure StringStructure => http://msdn2.microsoft.com/fr-fr/library/ms646987.aspx
  StringStructure = Packed Record
    wLength      : WORD;
    wValueLength : WORD;
    wType        : WORD;
    szKey        : WideChar;                 //Nom de la variable de taille quelconque
    Padding      : WORD;
    Value        : WideChar;                 //Valeur de la variable de taille wValueLength exprimé en WideChar si wType = 1 sinon exprimé en Octets.
  End;

  //Déscription de la structure StringTable => http://msdn2.microsoft.com/fr-fr/library/ms646992.aspx
  StringTable = Packed Record
    wLength      : WORD;
    wValueLength : WORD;
    wType        : WORD;
    szKey        : Array[1..9] Of WideChar;   //8 Caractères du 'Microsoft Standard Language identifier' + Le 0 Terminateur (Ex '040C04E4').
    //Padding    : WORD;                      //Pas besoin de Padding pour être aligné sur des DWORD.
    Children     : StringStructure;
  End;
  PStringStructure = ^StringStructure;

  //Déscription de la structure StringFileInfo => http://msdn2.microsoft.com/fr-fr/library/ms646989.aspx
  StringFileInfo = Packed Record
    wLength      : WORD;
    wValueLength : WORD;
    wType        : WORD;
    szKey        : Array[1..15] Of WideChar;    //14 Caractères de 'StringFileInfo' + Le 0 Terminateur.
    //Padding    : WORD;                        //Pas besoin de Padding pour être aligné sur des DWORD.
    Children     : StringTable;
  End;

  //Déscription de la structure VS_VERSIONINFO => http://msdn2.microsoft.com/fr-fr/library/ms647001.aspx
  VS_VERSIONINFO = Packed Record
    wLength      : WORD;
    wValueLength : WORD;
    wType        : WORD;
    szKey        : Array[1..16] Of WideChar;    //15 Caractères de 'VS_VERSION_INFO' + Le 0 Terminateur.
    Padding1     : WORD;                        //Le Padding pour être aligné sur des DWORD.
    Value        : VS_FIXEDFILEINFO;
    //Padding2   : WORD;                        //Pas besoin de Padding pour être aligné sur des DWORD.
    Children     : WORD;
  End;

  TFileStreamString = Class( TFileStream )
    Function WriteStringLn( Texte : String ) : Integer;
  End;

var  hFicLog        : TFileStreamString;

Const
  AGRANDISSEMENT_MAX : Integer = 4 * Sizeof(WideChar) * 2;  //Passer de 8.0.743.0 à 8.0.743.65535 deux fois (Pour FileVersion et pour ProductVersion)

(*******************************************************************************************
  Méthode qui converti une valeur de chaine en un WORD avec des contrôles.
*******************************************************************************************)
Function ConvertirNumero( ValeurChaine : String; NomObjet : String ) : WORD;
Var
  Valeur : Integer;

Begin

  ValeurChaine := Trim(ValeurChaine);
  If ValeurChaine = '' Then
    Raise Exception.CreateFmt( 'La valeur du %s ne peut pas être vide.', [NomObjet] );

    //** Tenter de convertir la chaine en un entier **//
  Try
    Valeur := StrToInt( ValeurChaine );
  Except
    On Ex : Exception Do
    Begin
      Raise Exception.CreateFmt( 'La valeur du %s est incorrect.'#13#10'%s', [NomObjet,Ex.Message] );
    End;
  End;

    //** Contrôler les valeurs limites **//
  If Valeur < 0 Then
  Begin
    Raise Exception.CreateFmt( 'La valeur du %s ne peut pas être négative', [NomObjet] );
  End;

  If Valeur > High(Word) Then
  Begin
    Raise Exception.CreateFmt( 'La valeur du %s ne peut pas dépasser la valeur %d.', [NomObjet,High(Word)] );
  End;

  Result := Valeur;

End;

(*******************************************************************************************
  Méthode qui écrit une chaine dans un flux après lui avoir ajoutée RC.
*******************************************************************************************)
Function TFileStreamString.WriteStringLn( Texte : String ) : Integer;
Var
  TexteLn : String;
Begin
  TexteLn := Texte + #13#10;
  Result  := Write( PChar(TexteLn)^, Length(TexteLn) );
End;

(*******************************************************************************************
  Fonction permettant d'aligner un Pointeur sur des doubles mots.
*******************************************************************************************)
Function AlignerDWORD( p : Pointer ) : Pointer; Overload;
Begin
  Result := Pointer( Integer( (PChar(p) + Sizeof(DWORD) - 1) ) And Not (Sizeof(DWORD) - 1) );
End;

Function AlignerDWORD( p : Integer ) : Integer; Overload;
Begin
  Result := (p + Sizeof(DWORD) - 1) And Not (Sizeof(DWORD) - 1);
End;

function READTOKENPipe(var St: string; Sep: string): string;
var
  i: Integer;
begin
  i := Pos(Sep, St);
  if i <= 0 then
    i := Length(St) + 1;
  Result := Copy(St, 1, i - 1);
  Delete(St, 1, i + Length(Sep) - 1);
end;

(*******************************************************************************************
  Déscription de la structure String => http://msdn2.microsoft.com/fr-fr/library/ms646987.aspx
  Retourne le nombre d'octets d'agrandissement ou de Réduction eventuel.
*******************************************************************************************)
Function TraiterStringStructure( Var sStringStructure : StringStructure
                               ; Var sVS_FIXEDFILEINFO : VS_FIXEDFILEINFO; Const pFin : PChar
                               ; Var Resultat : String;  VersionProduit : String = '' ) : Integer;
Var
  Agrandissement : Integer;
  TaillePadding  : Integer;
  TailleValeur   : Integer;
  pNewValue      : PWideChar;
  NewValue       : WideString;
  //hFicLog        : TFileStreamString;
  pValue         : PWideChar;
  pKey           : PWideChar;

Begin

  Result := 0;

  //Contrôler qu'il s'agisse bien de chaînes de caractères.
  If sStringStructure.wType <> 1 Then
    Raise Exception.Create( 'Erreur de structure ''String''' );

  pKey := @sStringStructure.szKey;

  //On ne traite que la Valeur de FileVersion et de ProductVersion
  //If (pKey <> 'FileVersion') And (pKey <> 'ProductVersion') Then Exit;
  If (pKey = 'FileVersion') Then // Exit;
  begin
    //Recalculer la nouvelle valeur sous forme de chaîne de caractères.
    WideFmtStr( NewValue, '%d.%d.%d.%d', [HIWORD(sVS_FIXEDFILEINFO.dwFileVersionMS)
                                         ,LOWORD(sVS_FIXEDFILEINFO.dwFileVersionMS)
                                         ,HIWORD(sVS_FIXEDFILEINFO.dwFileVersionLS)
                                         ,LOWORD(sVS_FIXEDFILEINFO.dwFileVersionLS)] );

    pValue := AlignerDWORD( pKey + Length(pKey) + 1 );

    //retourner le résultat sous forme de chaine et de fichier XML.
    Resultat := 'Incrémentation de ' + pValue + ' vers ' + NewValue;

    If hFicLog <> nil Then
    Begin
     // Try
        //NomFichierLog := ChangeFileExt(Trim(NomFichierLog), '.xml' );
       // hFicLog       := TFileStreamString.Create( NomFichierLog, fmCreate );

        Try
          //hFicLog.WriteStringLn( '<?xml version="1.0" encoding="iso-8859-1"?>' );
          hFicLog.WriteStringLn( '<IncBuild>' );
          hFicLog.WriteStringLn( '  <NuméroAvant>' + pValue   + '</NuméroAvant>' );
          hFicLog.WriteStringLn( '  <NuméroAprès>' + NewValue + '</NuméroAprès>' );
          hFicLog.WriteStringLn( '</IncBuild>' );
        Except
        End;

     // Finally
     //   FreeAndNil( hFicLog );
     // End;
    End;
  End  // fin si FileVersion
  else
    If (pKey = 'ProductVersion') Then
    begin
      NewValue := VersionProduit;
      pValue := AlignerDWORD( pKey + Length(pKey) + 1 );
      Try

          hFicLog.WriteStringLn( '<ProductVersion>' );
          hFicLog.WriteStringLn( '  <Avant>' + pValue   + '</Avant>' );
          hFicLog.WriteStringLn( '  <Après>' + NewValue + '</Après>' );
          hFicLog.WriteStringLn( '</ProductVersion>' );
      Except
      End;

      //retourner le résultat sous forme de chaine et de fichier XML.
     //  Resultat := 'Incrémentation de ' + pValue + ' vers ' + NewValue;

    end
  else exit;

  TailleValeur   := (Length(NewValue) + 1) * Sizeof(WideChar);
  TaillePadding  := 0;
  Agrandissement := AlignerDWORD(TailleValeur) - AlignerDWORD(sStringStructure.wValueLength * Sizeof(WideChar));
  //If Abs(Agrandissement) > AGRANDISSEMENT_MAX Then
  //  Raise Exception.Create( 'Erreur de calcul de l''Agrandissement' );

  //Il est nécessaire d'agrandir pour pouvoir insérer les nouveaux caractères.
  If Agrandissement > 0 Then
  Begin

    Move( pValue^, Pointer(PChar(pValue) + Agrandissement)^, pFin - pValue );

    //Remplir la zone de Padding
    TaillePadding := AlignerDWORD(TailleValeur) - TailleValeur;
    If TaillePadding > 0 Then
      Move( #0#0#0#0, Pointer( PChar(pValue) + ((Length(NewValue)+1) * Sizeof(WideChar)) )^, TaillePadding );

  End

  //Il est nécessaire de réduire.
  Else If Agrandissement < 0 Then
  Begin

    pNewValue := Pointer( PChar(pValue) + Abs(Agrandissement) );
    Move( pNewValue^, pValue^ ,pFin - pValue );

  End;

  //Copier la nouvelle valeur
  pNewValue := @NewValue[1];
  Move( pNewValue^, pValue^, TailleValeur );

  sStringStructure.wValueLength := Length(NewValue) + 1;
  sStringStructure.wLength      := sStringStructure.wLength + Agrandissement;
  Result                        := Agrandissement;

End;


(*******************************************************************************************
  Déscription de la structure StringTable => http://msdn2.microsoft.com/fr-fr/library/ms646992.aspx
  Retourne le nombre d'octets d'agrandissement ou de Réduction eventuel.
*******************************************************************************************)
Function TraiterStringTable( Var sStringTable : StringTable; Var sVS_FIXEDFILEINFO : VS_FIXEDFILEINFO
                           ; Const pFin : PChar; Var Resultat : String;  VersionProduit : String = '' ) : Integer;
Var
  TailleStructure : Integer;
  Agrandissement  : Integer;
  TailleTableau   : Integer;
  pString         : PStringStructure;

Begin

  If sStringTable.wValueLength <> 0 Then
    Raise Exception.Create( 'Erreur sur la taille de la structure ''StringTable''' );

  //Ne PAS faire de contrôle sur le Type car BORLAND => 0 et MICROSOFT => 1 !!!!
  //If sStringTable.wType <> 0 Then
  //  Raise Exception.Create( 'Erreur sur le type de la structure ''StringTable''' );

  //Traiter le champ children de type Tableau de StringStructure
  Agrandissement := 0;
  TailleTableau  := sStringTable.wLength - (Sizeof(sStringTable) - Sizeof(StringStructure));  //Conserver la taille du tableau en Octets.
  pString        := @sStringTable.Children;
  Repeat

    TailleStructure := AlignerDWORD( pString.wLength ); //Conserver la Taille de la Structure avant Modifications.
    Agrandissement  := Agrandissement + TraiterStringStructure( pString^, sVS_FIXEDFILEINFO, pFin, Resultat, VersionProduit );

    //Passer à la Structure String Suivante (Le membre wLength a put être modifié par la fonction TraiterStringStructure())
    TailleTableau := TailleTableau - TailleStructure;
    pString       := AlignerDWORD( PChar(pString) + pString.wLength );

  Until TailleTableau <= 0;

  sStringTable.wLength := sStringTable.wLength + Agrandissement;
  Result := Agrandissement;

End;


(*******************************************************************************************
  Déscription de la structure StringFileInfo => http://msdn2.microsoft.com/fr-fr/library/ms646989.aspx
  Retourne le nombre d'octets d'agrandissement ou de Réduction eventuel.
*******************************************************************************************)
Function TraiterStringFileInfo( Var sStringFileInfo : StringFileInfo; Var sVS_FIXEDFILEINFO : VS_FIXEDFILEINFO
                              ; Const pFin : PChar; Var Resultat : String;  VersionProduit : String = ''  ) : Integer;
Var
  Agrandissement : Integer;

Begin

  If sStringFileInfo.wValueLength <> 0 Then
    Raise Exception.Create( 'Erreur sur la taille de la structure ''StringFileInfo''' );

  //Ne PAS faire de contrôle sur le Type car BORLAND => 0 et MICROSOFT => 1 !!!!
  //If sStringFileInfo.wType <> 1 Then
  //  Raise Exception.Create( 'Erreur sur le type de la structure ''StringFileInfo''' );

  //Contrôler la clé de contrôle.
  If WideCompareStr( sStringFileInfo.szKey, 'StringFileInfo' ) <> 0 Then
    Raise Exception.Create( 'Erreur sur la clé de la structure ''StringFileInfo''' );

  //Traiter le champ children de type StringTable
  Agrandissement := TraiterStringTable( StringTable((@sStringFileInfo.Children)^), sVS_FIXEDFILEINFO, pFin, Resultat,  VersionProduit );

  sStringFileInfo.wLength := sStringFileInfo.wLength + Agrandissement;
  Result := Agrandissement;

End;

(*******************************************************************************************
  Déscription de la structure VS_VERSIONINFO => http://msdn2.microsoft.com/fr-fr/library/ms647001.aspx
  Retourne le nombre d'octets d'agrandissement ou de Réduction eventuel.
*******************************************************************************************)
Function TraiterVS_VERSIONINFO( Var sVS_VERSIONINFO : VS_VERSIONINFO; Const pFin : PChar;
                                Var Resultat : String;  ValeurForcee : String = '';
                                BuildOuVersion : BUILD_OU_VERSION = BUILD_FORCE; VersionProduit : String = '' ) : Integer;
Var
  Agrandissement : Integer;
  VersionMajeur  : WORD;
  VerionMineure  : WORD;
  Release        : WORD;
  Build          : WORD;

Begin

  If sVS_VERSIONINFO.wValueLength < SizeOf(VS_FIXEDFILEINFO) Then
    Raise Exception.Create( 'Erreur sur la taille de la structure ''VS_VERSIONINFO''' );

  //Contrôler la clé de contrôle.
  If WideCompareStr( sVS_VERSIONINFO.szKey, 'VS_VERSION_INFO' ) <> 0 Then
    Raise Exception.Create( 'Erreur sur la clé de la structure ''VS_VERSIONINFO''' );

  If sVS_VERSIONINFO.wType <> 0 Then
    Raise Exception.Create( 'Erreur sur le type de la structure ''VS_VERSIONINFO''' );

  //Si on n'a pas de valeur forcée on relit les valeurs dans le .res
  If ValeurForcee = '' Then
  Begin

    //Déscription de la structure VS_FIXEDFILEINFO => http://msdn2.microsoft.com/fr-fr/library/ms646997.aspx
    VersionMajeur := HIWORD(sVS_VERSIONINFO.Value.dwFileVersionMS);
    VerionMineure := LOWORD(sVS_VERSIONINFO.Value.dwFileVersionMS);
    Release       := HIWORD(sVS_VERSIONINFO.Value.dwFileVersionLS);
    Build         := LOWORD(sVS_VERSIONINFO.Value.dwFileVersionLS);

    //Incrémenter le numéro de build
    If Build < High(WORD) Then
      Inc(Build)
    Else
      Build := 0;

  End
  Else
  Begin

    //On relit les valeur dans le fichier .res si la valeur forcée correspond seulement au build.
    If BuildOuVersion = BUILD_FORCE Then
    Begin

      //Déscription de la structure VS_FIXEDFILEINFO => http://msdn2.microsoft.com/fr-fr/library/ms646997.aspx
      VersionMajeur := HIWORD(sVS_VERSIONINFO.Value.dwFileVersionMS);
      VerionMineure := LOWORD(sVS_VERSIONINFO.Value.dwFileVersionMS);
      Release       := HIWORD(sVS_VERSIONINFO.Value.dwFileVersionLS);
      Build         := ConvertirNumero( ValeurForcee, 'Numéro de Build' );

    End
    Else
    Begin

      //sinon on force la valeur du numéro de version.
      VersionMajeur := ConvertirNumero( ReadTokenPipe( ValeurForcee , '.' ), 'Numéro de Version Majeure' );
      VerionMineure := ConvertirNumero( ReadTokenPipe( ValeurForcee , '.' ), 'Numéro de Version Mineure' );
      Release       := ConvertirNumero( ReadTokenPipe( ValeurForcee , '.' ), 'Numéro de Release' );
      Build         := ConvertirNumero( ReadTokenPipe( ValeurForcee , '.' ), 'Numéro de Build' );

    End;

  End;

  sVS_VERSIONINFO.Value.dwFileVersionMS := MakeLong( VerionMineure, VersionMajeur );
  sVS_VERSIONINFO.Value.dwFileVersionLS := MakeLong( Build, Release );

  //sVS_VERSIONINFO.Value.dwProductVersionMS := sVS_VERSIONINFO.Value.dwFileVersionMS;
  //sVS_VERSIONINFO.Value.dwProductVersionLS := sVS_VERSIONINFO.Value.dwFileVersionLS;

  //Traiter le champ children de type StrinFileInfo
  Agrandissement := TraiterStringFileInfo( StringFileInfo((@sVS_VERSIONINFO.Children)^), sVS_VERSIONINFO.Value, pFin, Resultat, VersionProduit );

  sVS_VERSIONINFO.wLength := sVS_VERSIONINFO.wLength + Agrandissement;
  Result := Agrandissement;

End;

(*******************************************************************************************
   Cette Fonction Incrémente le Numéro de Build du fichier .res passer en paramètre.
   On peut forcer la version à enregistrer voir seulement le numéro de Build.
   Le résultat peut aussi être stocké dans un fichier XML.
   Paramètres :
      - CheminRes      : Chemin du Fichier .res à modifier.
      - CheminXml      : Chemin du Fichier XML de résultat.
      - ValeurForcee   : Valeur forcée pour la version ou le Build selon le paramètre BuildOuVersion
      - BuildOuVersion : Paramètre qui précise le contenu de ValeurForcee.
********************************************************************************************)
Function IncrementerBuild( CheminRes : String; CheminXml : String = ''; ValeurForcee : String = '';
                           BuildOuVersion : BUILD_OU_VERSION = BUILD_FORCE; VersionProduit : String = '' ) : String;

Var
  TailleRessource : Integer;
  Agrandissement  : Integer;
  TailleFichier   : Integer;
  CheminSource    : String;
  CheminDestin    : String;
  NomFichierLog   : String;
  TailleALire     : Integer;
  hFicSource      : TFileStream;
  hFicDestin      : TFileStream;
  HeaderSize      : DWORD;
  DataSize        : DWORD;
  Ressource       : Pointer;
  CodeType        : WORD;
  Control         : WORD;
  NomType         : PWideChar;
  Entete          : PTDWORD;

Begin

  Result := '';
  CheminSource := ChangeFileExt(CheminRes, '.res' );
  CheminDestin := ChangeFileExt(CheminRes, '.~res' );

  Try   //Toujours tenter de nettoyer le fichier temporaire.  ==> DeleteFile( CheminDestin );
    hFicLog := nil;
    if CheminXml <>'' then
    begin
      NomFichierLog := ChangeFileExt(Trim(CheminXml), '.xml' );
      hFicLog       := TFileStreamString.Create( NomFichierLog, fmCreate );
      hFicLog.WriteStringLn( '<?xml version="1.0" encoding="iso-8859-1"?>' );
    end;
    Try
      hFicSource := TFileStream.Create( CheminSource, fmOpenRead );
    Except
      //Les projets ne comporte pas toujours un fichier.res comme les applications de type CONSOLE par exemple.
      On EFOpenError Do
      Begin
        Result := 'Impossible d''ouvrir le fichier ' + CheminSource + '.'#13#10'Opération abandonnée';
        Exit;
      End;
    End;

    Try   //Bloc ==> FreeAndNil( hFicSource );

      hFicDestin := TFileStream.Create( CheminDestin, fmCreate );

      Try   //Bloc ==> FreeAndNil( hFicDestin );

        TailleFichier := hFicSource.Size;

        //Parcours des ressources du Fichier .RES
        Repeat

          //Lire la taille des données de la ressource.
          TailleALire := SizeOf(DataSize);
          Try
            hFicSource.ReadBuffer( DataSize, TailleALire );
          Except
            Raise Exception.CreateFmt( 'Erreur de lecture du fichier %s.', [CheminSource] );
          End;
          TailleFichier := TailleFichier - TailleALire;

          //Lire la taille de l'Entête.
          TailleALire := SizeOf(HeaderSize);
          Try
            hFicSource.ReadBuffer( HeaderSize, TailleALire );
          Except
            Raise Exception.CreateFmt( 'Erreur de lecture du fichier %s.', [CheminSource] );
          End;
          TailleFichier := TailleFichier - TailleALire;
          If HeaderSize <= (3 * SizeOf(DWORD)) Then
            Raise Exception.CreateFmt( 'Erreur de lecture du fichier %s.#10#13 Format Incorrect, Entête Vide', [CheminSource] );

          //Lire le reste de l'Entête.
          Entete := PTDWORD(AllocMem( HeaderSize ));
          
          Try   //Bloc ==> If Entete <> Nil Then FreeMem( Entete );
            Entete[0] := DataSize;
            Entete[1] := HeaderSize;

            TailleALire := HeaderSize - SizeOf(DataSize) - SizeOf(HeaderSize);
            Try
              hFicSource.ReadBuffer( Entete[2], TailleALire );
            Except
              Raise Exception.CreateFmt( 'Erreur de lecture du fichier %s.', [CheminSource] );
            End;
            TailleFichier := TailleFichier - TailleALire;

            //Lire le 1er WORD du type.
            Control := LOWORD(Entete[2]);
            If Control = $FFFF Then
            Begin

              //Lire le Code du type.
              CodeType := HIWORD(Entete[2]);

            End
            Else
            Begin

              NomType := PWideChar(@Entete[2]);

              //Convertir le nom en Code
              If NomType      = 'RT_VERSION'      Then CodeType := DWORD(RT_VERSION)
              Else If NomType = 'RT_CURSOR'       Then CodeType := DWORD(RT_CURSOR)
              Else If NomType = 'RT_BITMAP'       Then CodeType := DWORD(RT_BITMAP)
              Else If NomType = 'RT_ICON'         Then CodeType := DWORD(RT_ICON)
              Else If NomType = 'RT_MENU'         Then CodeType := DWORD(RT_MENU)
              Else If NomType = 'RT_DIALOG'       Then CodeType := DWORD(RT_DIALOG)
              Else If NomType = 'RT_STRING'       Then CodeType := DWORD(RT_STRING)
              Else If NomType = 'RT_FONTDIR'      Then CodeType := DWORD(RT_FONTDIR)
              Else If NomType = 'RT_FONT'         Then CodeType := DWORD(RT_FONT)
              Else If NomType = 'RT_ACCELERATOR'  Then CodeType := DWORD(RT_ACCELERATOR)
              Else If NomType = 'RT_RCDATA'       Then CodeType := DWORD(RT_RCDATA)
              Else If NomType = 'RT_MESSAGETABLE' Then CodeType := DWORD(RT_MESSAGETABLE)
              Else If NomType = 'RT_GROUP_CURSOR' Then CodeType := DWORD(RT_GROUP_CURSOR)
              Else If NomType = 'RT_GROUP_ICON'   Then CodeType := DWORD(RT_GROUP_ICON)
              Else If NomType = 'RT_DLGINCLUDE'   Then CodeType := DWORD(RT_DLGINCLUDE)
              Else If NomType = 'RT_PLUGPLAY'     Then CodeType := DWORD(RT_PLUGPLAY)
              Else If NomType = 'RT_VXD'          Then CodeType := DWORD(RT_VXD)
              Else If NomType = 'RT_ANICURSOR'    Then CodeType := DWORD(RT_ANICURSOR)
              Else If NomType = 'RT_ANIICON'      Then CodeType := DWORD(RT_ANIICON)
              Else CodeType := 0;

            End;

            //Les Entêtes de Ressources sont alignés sur des Doubles Mots.
            TailleRessource := AlignerDWORD( DataSize );

            //Lire le reste de la ressource.
            Try   //Bloc ==> If Ressource <> Nil Then FreeMem( Ressource );

              If TailleRessource > 0 Then
              Begin

                //On alloue AGRANDISSEMENT_MAX caractères supplémentaires pour tenir compte d'un eventuel agrandissement.
                Ressource := AllocMem( TailleRessource + (AGRANDISSEMENT_MAX * SizeOf(WideChar)) );
                Try
                  hFicSource.ReadBuffer( Ressource^, TailleRessource );
                Except
                  Raise Exception.CreateFmt( 'Erreur de lecture du fichier %s.', [CheminSource] );
                End;
                TailleFichier := TailleFichier - TailleRessource;

                If MakeIntResource(CodeType) = RT_VERSION Then
                Begin

                  //Modifier la ressource VS_VERSIONINFO
                  Agrandissement := TraiterVS_VERSIONINFO( VS_VERSIONINFO(Ressource^)
                                                         , PChar(PChar(Ressource) + TailleRessource)
                                                         , Result,  ValeurForcee, BuildOuVersion, VersionProduit );

                  TailleRessource := TailleRessource + Agrandissement;
                  Entete[0]       := WORD( Integer(Entete[0]) + Agrandissement );
                End;

              End;

              //Ranger l'Entête éventuellement modifié dans le fichier de destination.
              Try
                hFicDestin.WriteBuffer( Entete^, HeaderSize );
              Except
                Raise Exception.CreateFmt( 'Erreur d''écriture dans le fichier %s.', [CheminDestin] );
              End;

              //Ranger la ressource éventuellement modifiée dans le fichier de destination.
              If TailleRessource > 0 Then
                Begin
                Try
                  hFicDestin.WriteBuffer( Ressource^, TailleRessource );
                Except
                  Raise Exception.CreateFmt( 'Erreur d''écriture dans le fichier %s.', [CheminDestin] );
                End;
              End;

            Finally
              If Ressource <> Nil Then FreeMem( Ressource );
              Ressource := Nil;
            End;

          Finally
            If Entete <> Nil Then FreeMem( Entete );
            Entete := Nil;
          End;

        Until TailleFichier <= 0;

      Finally
        FreeAndNil( hFicDestin );
      End;

    Finally
      FreeAndNil( hFicSource );
    End;

    //Après la création du fichier Temporaire
    If Not DeleteFile( CheminSource ) Then
      Raise Exception.CreateFmt( 'Impossible de détruire le fichier ''%s.''', [CheminSource] );

    If Not RenameFile( CheminDestin, CheminSource ) Then
      Raise Exception.CreateFmt( 'Impossible de renommer le fichier de ''%s'' en ''%s.''', [CheminDestin,CheminSource] );

  Finally
    //Toujours tenter de nettoyer le fichier temporaire.
    DeleteFile( CheminDestin );
    FreeAndNil( hFicLog );
  End;

End;

End.

