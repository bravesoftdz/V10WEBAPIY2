{***********UNITE*************************************************
Auteur  ...... : Y. Peuillon
Créé le ...... : 04/11/2003
Modifié le ... : 04/11/2003
Description .. : Outil générique de transformation de données
Mots clefs ... : TRANSFORMATION;GPAO
*****************************************************************}
unit AssistTransformeGPAO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HPanel, assist, StdCtrls, Mask, Hctrls, HStatus, ComCtrls, HSysMenu, hmsgbox,
  ExtCtrls, HTB97, Hent1;

type
  TTTransformeGPAO = class(TFAssist)
    Lancement: TTabSheet;
    TabSheet1: TTabSheet;
    HLabel8: THLabel;
    FichIn: THCritMaskEdit;
    FichOut: THCritMaskEdit;
    Panel1: TPanel;
    TITREASSISTANT: THLabel;
    CompteRendu: TMemo;
    bRecup: TToolbarButton97;
    Label1: TLabel;
    LabelIdEntete: TLabel;
    LabelIdLigne: TLabel;
    OpenFichGPAO: TOpenDialog;
    OpenFichOut: TOpenDialog;
    LU: TLabel;
    Traitees: TLabel;
    Erreur: TLabel;
    NbreLu: TEdit;
    NbreTraite: TEdit;
    NbreRejet: TEdit;
    NbreEcrit: TEdit;
    IdEntete: TEdit;
    IdLigne: TEdit;
    IdOut: TEdit;
    LabelIdOut: TLabel;
    LabelFichOut: TLabel;
    Label6: TLabel;
    LabelTransfo: TLabel;
    TransfoType: THValComboBox;
    Bevel1: TBevel;
    Bevel2: TBevel;
    LabelFichIn: TLabel;
    bSave: TButton;
    procedure bPrecedentClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bRecupClick(Sender: TObject);
    procedure FichInElipsisClick(Sender: TObject);
    procedure FichOutElipsisClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure TransfoTypeChange(Sender: TObject);
    procedure bSaveClick(Sender: TObject);
  private   { Déclarations privées }
    //
    // Gestion des compteurs d'enregistrements
    //
    CptEnregTot       : integer ;        // Compteur des enregistrements lus (à traiter ou pas)
    CptTrt            : integer ;        // Compteur des enregistrements traités
    CptErr            : integer ;        // Compteur des enregistrement en erreurs
    CptOut            : integer ;        // Compteur des enregistrement écrits
    TypeTransfo       : string ;
    procedure AfficheCompteur ;
    procedure AfficheDebutTraitement ;
    procedure RefreshCompteur ;
    procedure AfficheFinTraitement ;
    procedure EcrireInfo (Chtmp : String);
  public    { Déclarations publiques }
    function TransformationGPAO(MyTypeTransfo, FileInput, IdEntete, IdLigne, FileOutput, IdOut : String; Affichage : Boolean) : Boolean;

  end;


Procedure Assist_TransformeGPAO;


implementation

{$R *.DFM}
uses
  {$IFDEF EAGLCLIENT}
  UtileAGL,
  {$ELSE}
  MajTable,
  {$ENDIF}
  UtilGPAO, Utob, ParamSoc, dbtables, UIutil;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Y. Peuillon
Créé le ...... : 24/07/2001
Modifié le ... : 10/11/2003
Description .. : Module d'import de données
Mots clefs ... : 
*****************************************************************}
Procedure Assist_TransformeGPAO;
Var PP : THPanel ;
    X : TTTransformeGPAO ;
begin
  X:=TTTransformeGPAO.Create(Application) ;

  // Initialisation de la provenance de la donnée
  PP:=FindInsidePanel ;
  if (PP=Nil)or true then
  begin
    Try
      X.ShowModal;
    Finally
      X.Free;
    end;
  end else
  begin
   InitInside(X,PP) ;
   X.Show ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Y. Peuillon
Créé le ...... : 17/10/2001
Modifié le ... : 10/11/2003
Description .. : Affichage des compteurs
Mots clefs ... : 
*****************************************************************}
procedure TTTransformeGPAO.AfficheCompteur ;
begin
  NbreLu.Text     := IntToStr (CptEnregTot) ;
  NbreTraite.Text := IntToStr (CptTrt) ;
  NbreRejet.Text  := IntToStr (CptErr) ;
  NbreEcrit.Text  := IntToStr (CptOut) ;
//  delay( 10 );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Y.Peuillon
Créé le ...... : 10/11/2003
Modifié le ... :   /  /    
Description .. : Détermine si on rafraîchit les compteurs
Mots clefs ... :
*****************************************************************}
procedure TTTransformeGPAO.RefreshCompteur ;
var
  LePointDuTraitement : integer;
begin
  if CptEnregTot < 10000 then
  begin
    LePointDuTraitement :=  CptEnregTot mod 1000;
    if LePointDuTraitement = 0 then AfficheCompteur ;
  end else
  begin
    LePointDuTraitement :=  CptEnregTot mod 5000;
    if LePointDuTraitement = 0 then AfficheCompteur ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Y. Peuillon
Créé le ...... : 05/11/2003
Modifié le ... : 10/11/2003
Description .. : Affichage des infos de début de traitement
Mots clefs ... :
*****************************************************************}
procedure TTTransformeGPAO.AfficheDebutTraitement ;
begin
  EcrireInfo ('  ');
  EcrireInfo (TraduireMemoire(' Début de la transformation des informations le ') + DateToStr(Date) + TraduireMemoire(' à ') + TimeToStr(Time));
  EcrireInfo (TraduireMemoire(' Fichier lu : ') + FichIn.Text);
  if (TypeTransfo = 'RGP') and (TypeTransfo = 'AID') then
    EcrireInfo (TraduireMemoire(' Fichier écrit : ') + FichOut.Text);
  EcrireInfo ('  ');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Y. Peuillon
Créé le ...... : 17/10/2001
Modifié le ... : 10/11/2003
Description .. : Affichage des infos (compteurs) en fin de traitement
Mots clefs ... :
*****************************************************************}
procedure TTTransformeGPAO.AfficheFinTraitement ;
begin
  EcrireInfo ('  ');
  EcrireInfo(TraduireMemoire(' Fin de la transformation des informations le ') + DateToStr(Date) + TraduireMemoire(' à ') + TimeToStr(Time));

  if CptEnregTot > 1 then EcrireInfo ('   -> ' + IntToStr (CptEnregTot) + TraduireMemoire(' enregistrements lus '))
  else                    EcrireInfo ('   -> ' + IntToStr (CptEnregTot) + TraduireMemoire(' enregistrement lu'));

  if CptTrt   > 1    then EcrireInfo ('   -> ' + IntToStr (CptTrt) + TraduireMemoire(' enregistrements traités '))
  else                    EcrireInfo ('   -> ' + IntToStr (CptTrt) + TraduireMemoire(' enregistrement traité '));

  if CptErr   > 1    then EcrireInfo ('   -> ' + IntToStr (CptErr) + TraduireMemoire(' enregistrements non repris '))
  else if CptErr = 1 then EcrireInfo ('   -> ' + IntToStr (CptErr) + TraduireMemoire(' enregistrement non repris '))
  else if CptErr = 0 then EcrireInfo (TraduireMemoire('   -> Aucun enregistrement non repris '));

  if CptOut > 1      then EcrireInfo ('   -> ' + IntToStr (CptOut) + TraduireMemoire(' enregistrements écrits '))
  else if CptOut=1   then EcrireInfo ('   -> ' + IntToStr (CptOut) + TraduireMemoire(' enregistrement écrit '))
  else if CptOut=0   then EcrireInfo (TraduireMemoire('   -> Aucun enregistrement écrit '));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 05/11/2003
Description .. : Lancement de la transformation
Mots clefs ... :
*****************************************************************}
procedure TTTransformeGPAO.bRecupClick(Sender: TObject);
begin
inherited;
  TransformationGPAO( TypeTransfo, FichIn.Text, IdEntete.Text, IdLigne.Text, FichOut.Text, IdOut.Text, True);
end;


procedure TTTransformeGPAO.bSuivantClick(Sender: TObject);
begin
inherited;
  if (P.ActivePage.PageIndex = P.PageCount - 1) then bFin.enabled := True;
  bSave.Enabled := False;
end;

procedure TTTransformeGPAO.bPrecedentClick(Sender: TObject);
begin
inherited;
  bFin.enabled := False;
  bSave.Enabled := True;
end;

procedure TTTransformeGPAO.FichInElipsisClick(Sender: TObject);
begin
  inherited;
  OpenFichGPAO.InitialDir:=GetParamSoc ('SO_GCREPORLI');
  FichIn.DataType     := 'FILE' ;
  FichIn.Enabled      := True ;
  if OpenFichGPAO.execute then FichIn.text:=OpenFichGPAO.FileName;
end;


procedure TTTransformeGPAO.FichOutElipsisClick(Sender: TObject);
begin
  inherited;
  OpenFichOut.InitialDir:=GetParamSoc ('SO_GCREPORLI');
  FichOut.DataType     := 'FILE' ;
  FichOut.Enabled      := True ;
  if OpenFichOut.execute then FichOut.text:=OpenFichOut.FileName;
end;

procedure TTTransformeGPAO.bFinClick(Sender: TObject);
begin
  inherited;
  if not isInside(Self) then close ;
end;

procedure TTTransformeGPAO.FormShow(Sender: TObject);
begin
  inherited;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;

  FichIn.Text  := GetParamSoc ('SO_GCREPORLI');
  FichOut.Text := GetParamSoc ('SO_GCREPORLI');

  TransfoTypeChange( nil );
end;



procedure TTTransformeGPAO.FormDestroy(Sender: TObject);
begin
  inherited ;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

procedure TTTransformeGPAO.EcrireInfo(Chtmp : String);
begin
  CompteRendu.lines.add( Chtmp );
end;

procedure TTTransformeGPAO.TransfoTypeChange(Sender: TObject);
var
  Q           : TQUERY ;
  Tob_Transfo : TOB    ;
  SQL         : String ;
begin
  inherited;

  TypeTransfo := TransfoType.Value;

  LabelIdEntete.Enabled := true;
  IdEntete.Enabled := true;
  LabelIdLigne.Enabled := true;
  IdLigne.Enabled := true;
  LabelFichOut.Enabled := true;
  FichOut.Enabled := true;
  LabelIdOut.Enabled := true;
  IdOut.Enabled := true;

  SQL := 'SELECT * FROM REPRISETFGPAO WHERE GRT_TYPETRANSFO="'+TypeTransfo+'"' ;
  Q   := OpenSQL (SQL ,True) ;
  Tob_Transfo := TOB.CREATE ('REPRISETFGPAO', NIL, -1);
  Tob_Transfo.SelectDB('',Q);
  Ferme(Q);
  IdEntete.Text := Tob_Transfo.GetValue('GRT_IDENTETE');
  IdLigne.Text  := Tob_Transfo.GetValue('GRT_IDLIGNE');
  IdOut.Text    := Tob_Transfo.GetValue('GRT_IDOUT');
  FreeAndNil(Tob_Transfo );

  if TypeTransfo = 'AID' then
  begin
    LabelIdEntete.Enabled := false;
    IdEntete.Enabled := false;
  end else if TypeTransfo = 'TRI' then
  begin
    LabelIdLigne.Enabled := false;
    IdLigne.Enabled := false;
    LabelFichOut.Enabled := false;
    FichOut.Enabled := false;
    LabelIdOut.Enabled := false;
    IdOut.Enabled := false;
  end;

end;

function TTTransformeGPAO.TransformationGPAO(MyTypeTransfo, FileInput, IdEntete, IdLigne, FileOutput, IdOut : String; Affichage : Boolean) : Boolean;
var
  FileIn, FileOut : TextFile;
  LnIn, EnteteOut, SIdOut, SIdOutSav, StID, TmpPath, mId : String;
  SIdEntete, SIdLigne, StTmp : PChar;
  APath:array[0..MAX_PATH] of Char;
  ListeFileId : TStringList;
  pId, i : integer;
begin
  Result := False;
  if Affichage then AfficheDebutTraitement;

  // Initialisation des compteurs
  CptEnregTot := 0;
  CptTrt := 0;
  CptErr := 0;
  CptOut := 0;

  if (MyTypeTransfo = 'RGP') or (MyTypeTransfo = 'AID') then
  begin
    if FileExists( FileInput ) then
    begin
        // Initialisation des Id
        EnteteOut := '';
        SIdEntete := PChar(IdEntete);
        SIdLigne := PChar(IdLigne);

        // Initialisation de la barre de progression
        InitMove( GetNbEnregInFile( FileInput ), '');

        // Si le fichier de sortie n'est pas renseigné
        if Trim(FileOutput) = '' then
        begin
          FileOutput := FileInput;
          FileInput := ChangeFileExt(FileOutput, '.OLD');
          i := 1;
          while FileExists( FileInput ) do
          begin
            FileInput := ChangeFileExt(FileOutput, '.OLD' + IntToStr(i));
            Inc( i );
          end;
          RenameFile( FileOutput, FileInput );
        end;

        // Ouverture du fichier en entrée
        AssignFile(FileIn, FileInput);
        Reset( FileIn );
        try
          // Ouverture du fichier en sortie
          AssignFile(FileOut, FileOutput);
          Rewrite( FileOut );
          try
            // Parcours du fichier en entrée
            while not SeekEof(FileIn) do
            begin
              Readln( FileIn, LnIn);
              Inc( CptEnregTot );
              if (Length(SIdEntete)<>0) and (strlcomp(SIdEntete, PChar(LnIn), StrLen(SIdEntete)) = 0) then
              begin // enregistrment entete
                EnteteOut := LnIn;
                Inc( CptTrt );
              end
              else if strlcomp(SIdLigne, PChar(LnIn), StrLen(SIdLigne)) = 0 then
              begin // enregistrement ligne
                SIdOutSav := IdOut;
                while SIdOutSav <> '' do
                begin
                  SIdOut := ReadTokenSt(SIdOutSav);
                  Writeln( FileOut, SIdOut + 'C1 ' + EnteteOut + LnIn );
                  Inc( CptOut );
                end;
                Inc( CptTrt );
              end
              else // autre enregistrement
                Inc( CptErr );

              // Rafraîchissement de l'affichage compteurs
              RefreshCompteur;
              // Augmentation de la barre de progression
              MoveCur( False );
            end;
            AfficheCompteur;
            AfficheFinTraitement;
          finally
            // Fermeture des fichiers en entrée et sortie
            CloseFile( FileOut);
          end;
        finally
          CloseFile( FileIn);
          FiniMove;
        end;
        Result := True;
    end else EcrireInfo (TraduireMemoire('Le Fichier ') + FichIn.Text + TraduireMemoire(' n''existe pas'))
  end
  else if MyTypeTransfo = 'TRI' then
  begin
    if FileExists( FileInput ) then
    begin
      // Initialisation de la barre de progression
      InitMove( GetNbEnregInFile( FileInput )* 2, '');

      // Traite chaque identifiant --> un fichier temporaire / ID
      if Affichage then EcrireInfo (TraduireMemoire(' Génération des fichiers temporaires...'));
      GetTempPath(MAX_PATH,APath);
      TmpPath := IncludeTrailingBackSlash(APath);

      // Ouverture du fichier en entrée
      AssignFile(FileIn, FileInput);
      Reset( FileIn );
      try
        ListeFileId := TStringList.Create;

        // Parcours du fichier en entrée
        while not SeekEof(FileIn) do
        begin
          Readln( FileIn, LnIn);
          mId := Copy( LnIn, 0, 3);
          Inc( CptEnregTot );
          pId := ListeFileId.IndexOfName( mId );
          if pId <> -1 then
          begin
            AssignFile(FileOut, TmpPath + mId + '.tmp');
            Append( FileOut );
          end
          else
          begin
            AssignFile(FileOut, TmpPath + mId + '.tmp');
            Rewrite( FileOut );
            ListeFileID.Add( mId + '=' + TmpPath + mId + '.tmp' );
          end;
          try
            Writeln( FileOut, LnIn );
            Inc( CptTrt );
          finally
            CloseFile( FileOut);
          end;

          // Rafraîchissement de l'affichage compteurs
          if Affichage then RefreshCompteur;
          // Augmentation de la barre de progression
          MoveCur( False );
        end;
      finally
        CloseFile( FileIn);
      end;
      if Affichage then
      begin
        EcrireInfo (TraduireMemoire(' Fin de génération des fichiers temporaires '));
        EcrireInfo (TraduireMemoire(' '));
        EcrireInfo (TraduireMemoire(' Reconstruction du fichier final... '));
      end;
      // Reconstruction du fichier trié
      StID := IdEntete;
      // Ouverture du fichier en entrée
      AssignFile(FileOut, FileInput);
      Rewrite( FileOut );
      try
        while StID<>'' do
        begin
          StTmp := PChar(ReadTokenSt(StID));

          // Ouverture d'un fichier temporaire
          if FileExists( TmpPath + StTmp + '.tmp' ) then
          begin
            AssignFile(FileIn, TmpPath + StTmp + '.tmp');
            Reset( FileIn );
            try
              // Parcours du fichier en entrée
              while not SeekEof(FileIn) do
              begin
                Readln( FileIn, LnIn);
                Writeln( FileOut, LnIn );
                Inc( CptEnregTot );
                Inc( CptOut );
                // Rafraîchissement de l'affichage compteurs
                if Affichage then RefreshCompteur;
                // Augmentation de la barre de progression
                MoveCur( False );
              end;
            finally
              // Fermeture du fichier en entrée
              CloseFile( FileIn);
            end;
            DeleteFIle(TmpPath + StTmp + '.tmp');
          end;
        end;
        for i := 0 to ListeFileId.Count-1 do
        begin
          mId := ListeFileId.Names[i];
          // Ouverture d'un fichier temporaire
          if FileExists( TmpPath + mId + '.tmp' ) then
          begin
            AssignFile(FileIn, TmpPath + mId + '.tmp');
            Reset( FileIn );
            try
              // Parcours du fichier en entrée
              while not SeekEof(FileIn) do
              begin
                Readln( FileIn, LnIn);
                Writeln( FileOut, LnIn );
                Inc( CptEnregTot );
                Inc( CptOut );
                // Rafraîchissement de l'affichage compteurs
                if Affichage then RefreshCompteur;
                // Augmentation de la barre de progression
                MoveCur( False );
              end;
            finally
              // Fermeture du fichier en entrée
              CloseFile( FileIn);
            end;
            DeleteFIle(TmpPath + mId + '.tmp');
          end;
        end;
      finally
        // Fermeture du fichier en sortie
        CloseFile( FileOut);
        ListeFileId.Free;
        FiniMove;
      end;
      if Affichage then AfficheCompteur;
      if Affichage then AfficheFinTraitement;
      Result := True;
    end else EcrireInfo (TraduireMemoire('Le Fichier ') + FileInput + TraduireMemoire(' n''existe pas'))
  end;
end;

procedure TTTransformeGPAO.bSaveClick(Sender: TObject);
var
  Q           : TQUERY ;
  Tob_Transfo : TOB    ;
  SQL         : String ;
begin
  inherited;
  SQL := 'SELECT * FROM REPRISETFGPAO WHERE GRT_TYPETRANSFO="'+TypeTransfo+'"' ;
  Q   := OpenSQL (SQL ,True) ;
  Tob_Transfo := TOB.CREATE ('REPRISETFGPAO', NIL, -1);
  Tob_Transfo.SelectDB('',Q);
  Ferme(Q);
  Tob_Transfo.PutValue('GRT_TYPETRANSFO', TypeTransfo);
  Tob_Transfo.PutValue('GRT_IDENTETE', IdEntete.Text);
  Tob_Transfo.PutValue('GRT_IDLIGNE', IdLigne.Text);
  Tob_Transfo.PutValue('GRT_IDOUT', IdOut.Text);
  Tob_Transfo.InsertOrUpdateDB;
  FreeAndNil(Tob_Transfo );
end;

end.
