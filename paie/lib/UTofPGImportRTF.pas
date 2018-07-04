{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEXPORTRTF ()
Mots clefs ... : TOF;PGEXPORTRTF
*****************************************************************}
{
PT1   : 31/05/2007 JL V_72 FQ 14022 Contrôle de la langue
PT2   : 18/12/2007 VG V_80 Mauvaise alimentation du champ Nodossier en mode PCL
                           FQ N°15071
PT3   : 23/01/2008 VG V_80 Sortie de la combo de la grille - FQ N°13707                           
}
Unit UTofPGImportRTF ;

Interface

Uses StdCtrls,
     Controls, 
     Classes,
     Graphics,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, 
{$else}
     eMul, 

{$ENDIF}
     forms,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     PGOutils,
     UyFileSTD,
     HTB97,
     LookUp,
     Grids,
     paramsoc,
     FileCtrl,
     ed_tools,
     HSysMenu,
     Vierge,
     P5Util,
     UTOF,
     PgOutilsGrids;
Const
ColFichier = 0;
ColExtension = 1;
ColNom = 2;
ColCrit1 = 3;
ColCrit2 = 4;
ColEnBase = 5;
ColPredef = 6;
ColImport = 7;
ColConfidentiel = 8;
ColSalarie = 9;
ColSauvFic = 10;


Type
  TOF_PGIMPORTRTF = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad; override ;
    procedure OnClose  ; override ;
    private
    GFichier : THGrid;
    StChampGrid : String;
    TobFichiers,TF : Tob;
    procedure BImportClick(Sender : TObject);
    procedure RecupDoc(Sender : TObject);
    procedure GrilleElipsisClick (Sender : TObject);
    procedure SelectionImport(Sender : Tobject);
    procedure OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure VerifPresence(Ligne : Integer);
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure AccesChemin (Sender : TObject);
    procedure SuppressionDonnees(NomFic,Langue,Predef,NoDossier : String);
    procedure ColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GFClick (Sender: TObject);
    Procedure ChangeLangue(Sender : TObject);
  end ;


Implementation

procedure TOF_PGIMPORTRTF.OnClose ;
begin
     If TobFichiers <> Nil then FreeAndNil(TobFichiers);
end;

procedure TOF_PGIMPORTRTF.OnLoad;
begin
  GFIchier.ColFormats[ColCrit1] := 'CB=PGCATEGFICHIER';
  GFIchier.ColFormats[ColCrit2] := 'CB=PGNATUREFICHIER';
  GFIchier.ColFormats[ColPredef] := 'CB=PGPREDEFINI';
  GFIchier.ColTypes[ColEnBase] := 'B';
  GFIchier.ColFormats[ColEnBase] := IntToStr(Ord (csCheckBox));
  GFIchier.ColAligns[ColEnBase] := taCenter;
  GFIchier.ColTypes[ColConfidentiel] := 'B';
  GFIchier.ColFormats[ColConfidentiel] := IntToStr(Ord(csCheckBox));
  GFIchier.ColAligns[ColConfidentiel] := taCenter;
  GFIchier.ColTypes[ColImport] := 'B';
  GFIchier.ColFormats[ColImport] := IntToStr(Ord(csCheckBox));
  GFIchier.ColAligns[ColImport] := taCenter;
  GFichier.ColEditables[ColEnBase] := False;
  GFichier.ColEditables[ColImport] := False;
  GFichier.ColEditables[ColConfidentiel] := False;
  GFichier.ColWidths[ColSauvFic] := -1 ;
  GFichier.OnClick := GFClick;
  GFichier.ColWidths[ColFichier] := 200;
  GFichier.ColWidths[ColExtension] := 80;
end;

procedure TOF_PGIMPORTRTF.OnArgument (S : String ) ;
var BImport,BRec,BSelect : TTOolBarButton97;
    Combo : THValComboBox;
begin
  Inherited ;
  SetControlText('PAYS','FRA');
  BImport := TTOolBarButton97(GetControl('BValider'));
  If BImport <> Nil then BImport.OnClick := BImportClick;
  BRec := TTOolBarButton97(GetControl('BRECHERCHER'));
  If BRec <> Nil then BRec.OnClick := RecupDoc;
  BSelect := TTOolBarButton97(GetControl('BSelectAll'));
  If BSelect <> Nil then BSelect.OnClick := SelectionImport;
  GFichier := THGrid(GetControl('GRIDRESULT'));

  GFichier.OnColEnter := ColEnter;
  GFichier.OnCellExit := OnCellExit;
  GFichier.GetCellCanvas := GrilleGetCellCanvas;
  GFichier.OnElipsisClick := GrilleElipsisClick;
  GFichier.PostDrawCell := GrillePostDrawCell;

  Combo := THValComboBox(GetControl('TYPECHEMIN'));
  If Combo <> Nil then Combo.OnChange := AccesChemin;
  SetControlText('TYPECHEMIN','LIBRE');

  //on cache la gestion salarié
  SetControlVisible('DOUBLONSAL',false);
  GFichier.ColWidths[ColSalarie] := -1 ;

  Combo := THValComboBox(GetControl('PAYS'));
  If Combo <> Nil then Combo.OnChange := ChangeLangue;
  SetControlenabled('BValider',False);
  SetControlText('PREDEFINI','DOS');
  SetControlText('CRITERE2','');
end ;

procedure TOF_PGIMPORTRTF.OnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var Libelle : String;
begin
  If (ACol = ColPredef) or (ACol = ColCrit1) then VerifPresence(ARow);
  If (ACol = ColSalarie) and (GFichier.CellValues[ColSalarie,ARow]<>'')then
  begin
    GFichier.CellValues[ColPredef,ARow] := 'DOS';
    VerifPresence(ARow);
  end;
  If (ACol = ColPredef) and (GFichier.CellValues[ColPredef,ARow]<>'DOS') then GFichier.CellValues[ColSalarie,ARow] := '';
  if (ACol = ColNom) then
  begin
       Libelle := GFichier.CellValues[ColNom,ARow];
       If Length(Libelle) > 35 then
       begin
            PGIBox('Attention le libellé ne doit pas comporter plus de 35 caractères sinon il sera tronqué',Ecran.Caption);
            GFichier.Col := ACol;
            GFichier.Row := ARow;
       end;
  end;
end;

procedure TOF_PGIMPORTRTF.VerifPresence(Ligne : Integer);
var Crit1,Nom,Predef,NoDossier : String;
begin
  If Ligne = 0 then Exit;
  Predef := GFichier.CellValues[ColPredef,Ligne];
  Nom := GFichier.CellValues[ColFichier,Ligne] + GFichier.CellValues[ColExtension,Ligne];
  Crit1 := GFichier.CellValues[ColCrit1,Ligne];
  If Predef = 'DOS' then NoDossier := V_PGI.NoDossier
  else NoDossier := '000000';
  If ExisteSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_LANGUE="'+GetControlText('PAYS')+'" AND '+
  'YFS_CODEPRODUIT="PAIE" AND YFS_NOM="'+Nom+'" AND YFS_NODOSSIER="'+NoDossier+'" AND '+
  'YFS_CRIT1="'+Crit1+'" AND YFS_PREDEFINI="'+Predef+'"') then
  GFichier.CellValues[ColEnBase,Ligne] := 'X'
  else GFichier.CellValues[ColEnBase,Ligne] := '-';
end;

procedure TOF_PGIMPORTRTF.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin
  If (GFichier.CellValues[ColEnBase,ARow] = 'X') then Canvas.Font.Color := Clred;
end;

procedure TOF_PGIMPORTRTF.AccesChemin (Sender : TObject);
begin
     If getControlText('TYPECHEMIN') = 'LIBRE' then
     begin
          SetControlEnabled('TXTPATH',True);
          SetControlText('TXTPATH','');
     end
     else
     begin
          SetControlEnabled('TXTPATH',False);
          SetControlText('TXTPATH',GetParamSocSecur(GetControlText('TYPECHEMIN'), False ));
     end;
end;

procedure TOF_PGIMPORTRTF.BImportClick(Sender : TObject);
var
FilePath : String;
Crit1,Crit2,Nom,Libelle,NomFichier : String;
CodeRetour : Integer;
Langue,Predef : String;
Erreur,Ecraser,SuppFic,ErreurCegid : Boolean;
Q : TQuery;
i,NbImport : Integer;
NoDossier,Salarie,Confidentiel : String;
TobMaj : Tob;
HMTrad: THSystemMenu;
Extension,NomCourt : String;
begin
//DEBUT PT1
Langue:= GetControlText ('PAYS');
If (Langue='') then
   begin
   PGIBox ('Import impossible, vous devez renseigner la langue', Ecran.Caption);
   Exit;
   end;
If ((UpperCase (RechDom ('TTLANGUE', Langue, False))='FRANCE') OR
   (UpperCase (RechDom ('TTLANGUE', Langue, False))='FRANçAIS') OR
   (UpperCase (RechDom ('TTLANGUE', Langue, False))='FRANCAIS')) and
   (Langue<>'FRA') then
   begin
   PGIBox ('Import impossible, le pays France doit être associé au code FRA',
           Ecran.Caption);
   Exit;
   end;
//FIN PT1
{PT3
GFichier.Col:= 2;
GFichier.Row:= 1;
}
ValideComboDansGrid (GFichier, GFichier.Col, 2, 1);
//FIN PT3
SuppFic:= False;
if (GetCheckBoxState ('SUPPFIC')=CbChecked) then
   SuppFic:= True;

Ecraser:= GetCheckBoxState ('ECRASEMENT')=CbChecked;
TobFichiers.GetGridDetail (GFichier, GFichier.RowCount-1, '', StChampGrid);
Erreur:= False;
ErreurCegid:= False;
NbImport:= 0;
For i:= 0 to TobFichiers.Detail.Count-1 do
    begin
    If (TobFichiers.Detail[i].GetValue ('IMPORT')='X') then
       begin
       NbImport:= NbImport+1;
       If (TobFichiers.Detail[i].GetValue ('CRIT1')='') then
          begin
          Erreur:= True;
          GFichier.GotoRow (i+1);
          Break;
          end;
       if (TobFichiers.Detail[i].GetValue ('PREDEFINI')='') then
          begin
          Erreur:= True;
          GFichier.GotoRow (i+1);
          Break;
          end;
       end;
    If (not DroitCegPaie) then
       begin
       If (TobFichiers.Detail[i].GetValue ('IMPORT')='X') then
          begin
          If (TobFichiers.Detail[i].GetValue ('PREDEFINI')='CEG') then
             begin
             ErreurCegid:= True;
             GFichier.GotoRow (i+1);
             Break;
             end;
          end;
       end;
    end;
If Erreur then
   begin
   PGIInfo ('La colonne Catégorie doit être intégralement remplie pour les'+
            ' fichiers à importer', Ecran.Caption);
   exit;
   end;
If ErreurCegid then
   begin
   PGIInfo ('Vous ne pouvez pas importer de fichier CEGID', Ecran.Caption);
   exit;
   end;
If Suppfic then
   if (PGIAsk ('Les fichiers importés vont être détruits DEFINITIVEMENT du disque.'+#13#10+
               'Voulez vous continuer ?', Ecran.Caption)=Mrno) then
      exit;
If (NbImport=0) then
   PGIBox ('Aucun fichier sélectionné', Ecran.Caption);
InitMoveProgressForm (NIL, 'Import des fichiers', 'Veuillez patienter SVP ...',
                      NbImport, FALSE, TRUE);
For i:= 0 to TobFichiers.Detail.Count-1 do
    begin
    If (TobFichiers.Detail[i].GetValue ('IMPORT')<>'X') then
       Continue;
    If (Not Ecraser) and (TobFichiers.Detail[i].GetValue ('ENBASE')='X') then
       Continue;
    Crit2:= TobFichiers.Detail[i].GetValue ('CRIT2');
    Nom:= TobFichiers.Detail[i].GetValue ('NOMFIC')+
          TobFichiers.Detail[i].GetValue ('EXTENSION');
    NomFichier:= Nom;
    Extension:= Nom;
    NomCourt:= ReadTokenPipe (Extension, '.');
    Salarie:= TobFichiers.Detail[i].GetValue ('SALARIE');
    If (GetCheckBoxState('DOUBLONSAL')=CbChecked) then
       Nom:= Salarie+Nom;
    Predef:= TobFichiers.Detail[i].GetValue ('PREDEFINI');
    MoveCurProgressForm (Nom);
    FilePath:= GetControlText ('TXTPATH');
    Crit1:= TobFichiers.Detail[i].GetValue ('CRIT1');
    Libelle:= TobFichiers.Detail[i].GetValue ('LIBELLE');

    Confidentiel:= TobFichiers.Detail[i].GetValue ('CONFIDENTIEL');
    If (Predef='DOS') then
       NoDossier:= V_PGI.NoDossier
    else
       NoDossier:= '000000';
{
    If (Nom<>TobFichiers.Detail[i].GetValue('SAUVNOMFIC')) then
       SuppressionDonnees (Nom, Langue, Predef, NoDossier);
}
    NomFichier:= TobFichiers.Detail[i].GetValue ('SAUVNOMFIC');
    ExecuteSQL ('UPDATE YFILESTD SET'+
                ' YFS_CRIT2="'+Crit2+'" WHERE'+
                ' YFS_LANGUE="'+GetControlText('PAYS')+'" AND'+
                ' YFS_CODEPRODUIT="PAIE" AND'+
                ' YFS_NOM="'+Nom+'" AND'+
                ' YFS_NODOSSIER="'+NoDossier+'" AND'+
                ' YFS_CRIT1="'+Crit1+'" AND'+
                ' YFS_PREDEFINI="'+Predef+'"');
{PT2
    CodeRetour:= AGL_YFILESTD_IMPORT (FilePath+'\'+NomFichier, 'PAIE', Nom,
                                      Extension, Crit1, Crit2, Salarie, '', '',
                                      '-', '-', '-', '-', '-', Langue, Predef,
                                      Libelle);
}
    CodeRetour:= AGL_YFILESTD_IMPORT (FilePath+'\'+NomFichier, 'PAIE', Nom,
                                      Extension, Crit1, Crit2, Salarie, '', '',
                                      '-', '-', '-', '-', '-', Langue, Predef,
                                      Libelle, NoDossier);
//FIN PT2
    If (Ecraser) and (TobFichiers.Detail[i].GetValue ('IMPORT')='X') and
       (TobFichiers.Detail[i].GetValue ('ENBASE')='X') then
       begin
       TobMaj:= Tob.Create ('YFILESTD', Nil, -1);
       Q:= OpenSQl ('SELECT *'+
                    ' FROM YFILESTD WHERE'+
                    ' YFS_LANGUE="'+GetControlText ('PAYS')+'" AND'+
                    ' YFS_CODEPRODUIT="PAIE" AND'+
                    ' YFS_NOM="'+Nom+'" AND'+
                    ' YFS_NODOSSIER="" AND'+
                    ' YFS_CRIT1="'+Crit1+'" AND'+
                    ' YFS_PREDEFINI="'+Predef+'"', True);
       TobMaj.LoadDetailDB ('YFILESTD', '', '', Q, False);
       Ferme (Q);
       If (TobMaj.Detail.Count=1) then
          TobMaj.Detail[0].PutValue ('YFS_NODOSSIER', NoDossier);
       TobMaj.InsertOrUpdateDB (False);
       TobMaj.Free;
       ExecuteSQL ('DELETE FROM YFILESTD WHERE'+
                   ' YFS_LANGUE="'+GetControlText ('PAYS')+'" AND'+
                   ' YFS_CODEPRODUIT="PAIE" AND'+
                   ' YFS_NOM="'+Nom+'" AND'+
                   ' YFS_NODOSSIER="" AND'+
                   ' YFS_CRIT1="'+Crit1+'" AND'+
                   ' YFS_PREDEFINI="'+Predef+'"');
       end;

    If (CodeRetour=-1) and (Confidentiel='X') then
       ExecuteSQL ('UPDATE YFILESTD SET'+
                   ' YFS_CONFIDENTIEL="1" WHERE'+
                   ' YFS_LANGUE="'+GetControlText ('PAYS')+'" AND'+
                   ' YFS_CODEPRODUIT="PAIE" AND'+
                   ' YFS_NOM="'+Nom+'" AND'+
                   ' YFS_NODOSSIER="'+NoDossier+'" AND'+
                   ' YFS_CRIT1="'+Crit1+'" AND'+
                   ' YFS_PREDEFINI="'+Predef+'"');

    If (CodeRetour=-1) and (Confidentiel='-') and (Ecraser) and
       (TobFichiers.Detail[i].GetValue ('ENBASE')='X') then
       ExecuteSQL ('UPDATE YFILESTD SET'+
                   ' YFS_CONFIDENTIEL="0" WHERE'+
                   ' YFS_LANGUE="'+GetControlText ('PAYS')+'" AND'+
                   ' YFS_CODEPRODUIT="PAIE" AND'+
                   ' YFS_NOM="'+Nom+'" AND'+
                   ' YFS_NODOSSIER="'+NoDossier+'" AND'+
                   ' YFS_CRIT1="'+Crit1+'" AND'+
                   ' YFS_PREDEFINI="'+Predef+'"');

    if (SuppFic) and (CodeRetour=-1) then
       if FileExists (pchar (FilePath+'\'+NomFichier)) then
          DeleteFile (pchar (FilePath+'\'+NomFichier));
    end;
FiniMoveProgressForm;

If SuppFic then
   begin
   For i:= TobFichiers.Detail.Count-1 downto 0 do
       begin
       If (TobFichiers.Detail[i].GetValue('IMPORT')='X') then
          TobFichiers.Detail[i].Free;
       end;
   end;
TobFichiers.PutGridDetail (GFichier, False, False, StChampGrid, False);
For i:= 0 to GFichier.RowCount-1 do
    VerifPresence (i);

HMTrad.ResizeGridColumns (GFichier);
end;

procedure TOF_PGIMPORTRTF.RecupDoc(Sender : TObject);
var
  sr: TSearchRec;
  FileAttrs: Integer;
  stPath : string;
  i : Integer;
  Ext,ExtUtil,NomFichier,Extension,FichierBase,Crit1 : String;
  Crit2,Predef : String;
  TobExtension,T : Tob;
  Tous,LibTronque : Boolean;
  HMTrad: THSystemMenu;
  Const ChaineEntier = ['0'..'9'] ;
begin
if not (DirectoryExists(GetControlText('TXTPATH'))) then
begin
    PgiInfo('Attention, le répertoire de recherche ' +
      GetControlText('TXTPATH') + ' n''existe pas.',
      'Répertoire Introuvable');
      exit;
end;
Crit2 := GetControlText('CRITERE2');
Predef := GetControlText('PREDEFINI');
If (Predef = 'CEG') and not (DroitCegPaie) then
begin
     PGIBox('Vous ne pouvez pas importer de fichier CEGID',Ecran.Caption);
     exit;
end;
LibTronque := false;
For i := GFichier.rowCount - 1 downto 1 do GFichier.DeleteRow(i);
StChampGrid := 'NOMFIC;EXTENSION;LIBELLE;CRIT1;CRIT2;ENBASE;PREDEFINI;IMPORT;CONFIDENTIEL;SALARIE;SAUVNOMFIC';
FileAttrs := faAnyFile;
stPath := GetControlText('TXTPATH')+'\*.*';
If TobFichiers<> Nil then FreeAndNil(TobFichiers);
TobFichiers := Tob.Create('LesFichiers',Nil,-1);
TobExtension := Tob.Create('LesExtensions',Nil,-1);
ExtUtil := GetControlText('CMBEXT');
Ext := ReadTokenPipe(ExtUtil,';');
While Ext <> '' do
begin
     T := Tob.Create('Ext',TobExtension,-1);
     Ext := UpperCase(Ext);
     T.AddChampSupValeur('EXT',Ext);
     Ext := ReadTokenPipe(ExtUtil,';');
end;
If THMultiValComboBox(GetControl('CMBEXT')).Tous = True then Tous := True
Else Tous := False;
if FindFirst(stPath, FileAttrs, sr) = 0 then
  begin
        Extension := Sr.Name;
        NomFichier := ReadTokenPipe(Extension,'.');
        Fichierbase := Sr.Name;
        If Length(FichierBase) > 35 then
        begin
             LibTronque := True;
             NomFichier := Copy(NomFichier,0,31);
             FichierBase := Copy(FichierBase,0,31);
             FichierBase := FichierBase+'.'+Extension;
        end;
        Extension := UpperCase(Extension);
        If ((Extension <> '') and (Extension <> '.')) and ((TobExtension.FindFirst(['EXT'],[Extension],False) <> Nil) or (Tous)) then
        begin
            If Extension = 'DOC' then Crit1 := 'MAW'
            else If Extension = 'DOT' then Crit1 := 'MOW'
            else If Extension = 'XLS' then Crit1 := 'XLS'
            else If Extension = 'PDF' then Crit1 := 'PDF'
            else Crit1 := '';
            TF := Tob.Create('UnFichier',TobFichiers,-1);
            TF.AddChampSupValeur('NOMFIC', NomFichier);
            TF.AddChampSupValeur('EXTENSION', '.'+Extension);
            TF.AddChampSupValeur('SAUVNOMFIC', sr.Name);
            TF.AddChampSupValeur('LIBELLE', FichierBase);
            TF.AddChampSupValeur('CRIT1', Crit1);
            TF.AddChampSupValeur('CRIT2', Crit2);
            TF.AddChampSupValeur('CRIT3', '');
            TF.AddChampSupValeur('CRIT4', '');
            TF.AddChampSupValeur('CRIT5', '');
            TF.AddChampSupValeur('ENBASE', '-');
            TF.AddChampSupValeur('PREDEFINI', Predef);
            TF.AddChampSupValeur('IMPORT', 'X');
            TF.AddChampSupValeur('CONFIDENTIEL', '-');
            TF.AddChampSupValeur('SALARIE', '');
            TF.AddChampSupValeur('EXT', Extension);
        end;
        while (FindNext(sr) = 0) do
        begin
             Extension := Sr.Name;
             NomFichier := ReadTokenPipe(Extension,'.');
             Fichierbase := Sr.Name;
             If Length(FichierBase) > 35 then
             begin
                  LibTronque := True;
                  NomFichier := Copy(NomFichier,0,31);
                  FichierBase := Copy(FichierBase,0,31);
                  FichierBase := FichierBase+'.'+Extension;
             end;
             Extension := UpperCase(Extension);
             If ((Extension <> '') and (Extension <> '.')) and ((TobExtension.FindFirst(['EXT'],[Extension],False) <> Nil) or (Tous)) then
             begin
                If Extension = 'DOC' then Crit1 := 'MAW'
                else If Extension = 'DOT' then Crit1 := 'MOW'
                else If Extension = 'XLS' then Crit1 := 'XLS'
                else If Extension = 'PDF' then Crit1 := 'PDF'
                else Crit1 := '';
                TF := Tob.Create('UnFichier',TobFichiers,-1);
                TF.AddChampSupValeur('NOMFIC', NomFichier);
                TF.AddChampSupValeur('EXTENSION', '.'+Extension);
                TF.AddChampSupValeur('SAUVNOMFIC', sr.Name);
                TF.AddChampSupValeur('LIBELLE', FichierBase);
                TF.AddChampSupValeur('CRIT1', Crit1);
                TF.AddChampSupValeur('CRIT2', Crit2);
                TF.AddChampSupValeur('CRIT3', '');
                TF.AddChampSupValeur('CRIT4', '');
                TF.AddChampSupValeur('CRIT5', '');
                TF.AddChampSupValeur('ENBASE', '-');
                TF.AddChampSupValeur('PREDEFINI', Predef);
                TF.AddChampSupValeur('IMPORT', 'X');
                TF.AddChampSupValeur('CONFIDENTIEL', '-');
                TF.AddChampSupValeur('SALARIE', '');
                TF.AddChampSupValeur('EXT', Extension);
             end;
        end;
        sysutils.FindClose(sr);
  end;
If LibTronque then
begin
     PGIBox('Attention, certains nom de fichier comportant plus de 35 caractères ont été tronqué',Ecran.Caption);
end;
TobExtension.Free;
GFichier := THGrid(GetControl('GRIDRESULT'));
If TobFichiers.Detail.Count > 0 then SetControlenabled('BValider',True)
else SetControlenabled('BValider',False);
TobFichiers.PutGridDetail(GFichier,False,False,StChampGrid,False);
For i := 0 to GFichier.RowCount - 1 do
begin
     VerifPresence(i);
end;
//TobFichiers.Free;
 HMTrad.ResizeGridColumns (GFichier) ;
end;

procedure TOF_PGIMPORTRTF.SelectionImport(Sender : Tobject);
var Cocher : String;
    i : Integer;
begin
  If Sender = Nil then exit;
  If TToolBarButton97(Sender).Down = True then Cocher := 'X'
  else cocher := '-';
  For i := 1 to GFichier.RowCount - 1 do
  begin
    GFichier.CellValues[ColImport,i] := Cocher;
  end;
end;

procedure TOF_PGIMPORTRTF.ColEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
     If Ou = ColSalarie then THGrid(Sender).ElipsisButton := True
     Else THGrid(Sender).ElipsisButton := False;
end;


procedure TOF_PGIMPORTRTF.GrilleElipsisClick (Sender : TObject);
begin
    LookupList(THEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', '', 'PSA_SALARIE', True, -1);
end;

procedure TOF_PGIMPORTRTF.SuppressionDonnees(NomFic,Langue,Predef,NoDossier : String);
var Q : TQuery;
    Guid : String;
begin
     Guid := '';
     Q := OpenSQL('SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND '+
     'YFS_NOM="'+NomFic+'" AND YFS_LANGUE="'+Langue+'" AND YFS_PREDEFINI="'+Predef+'" AND YFS_NODOSSIER="'+NoDossier+'"',True);
     If Not Q.Eof then Guid := Q.FindField('YFS_FILEGUID').AsString;
     ferme(Q);
     ExecuteSQL('DELETE FROM YFILESTD WHERE YFS_FILEGUID="'+Guid+'"');
     ExecuteSQL('DELETE FROM NFILES WHERE NFI_FILEGUID="'+Guid+'"');
     ExecuteSQL('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID="'+Guid+'"');
end;

procedure TOF_PGIMPORTRTF.GrillePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  If (ACol = ColEnBase) then GridGriseCell(GFichier, Acol, Arow, Canvas);
end;

procedure TOF_PGIMPORTRTF.GFClick (Sender: TObject);
begin
if (GFichier.ColTypes[GFichier.col]='B') and (GFichier.col<>ColEnBase) then
   if GFichier.cells[GFichier.col,GFichier.Row]='X' then
   GFichier.cells[GFichier.col,GFichier.Row]:='-'
   else
   GFichier.cells[GFichier.col,GFichier.Row]:='X';
 //  NextPrevControl(TFVierge(Ecran));
end;

Procedure TOF_PGIMPORTRTF.ChangeLangue(Sender : TObject);
var i : Integer;
begin
     For i := 1 To GFichier.RowCount - 1 do
     begin
          VerifPresence(i);
     end;
end;


Initialization
  registerclasses ( [ TOF_PGIMPORTRTF ] ) ;
end.



