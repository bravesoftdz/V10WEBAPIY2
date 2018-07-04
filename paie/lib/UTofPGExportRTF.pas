{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEXPORTRTF ()
Mots clefs ... : TOF;PGEXPORTRTF
*****************************************************************}
{
PT1   : 11/04/2007 MF V_72 Modifs Mise en base des fichiers Ducs EDI
                           Pas d'export des fichiers DucsEDI==>Modif XX_WHERE
PT2   : 28/06/2007 VG V_72 Correction PT1
PT3   : 26/11/2007 VG V_80 Pas d'export des fichiers DADS-U
PT4   : 28/01/2008 VG V_80 Nouvelle question pour suppression du fichier
                           temporaire - FQ N°14382
}
Unit UTofPGExportRTF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     HDB,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     FE_Main,
{$else}
     eMul,
     MainEAGL,
{$ENDIF}
     forms,
     HQry,
     UTOB,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UYFileSTD,
     AGLUtilOle,
     UTobDebug,
     HTB97,
     UtilXLS,
     ShellAPI,
      Windows,
     Ed_Tools,
Menus,
     P5Util,
     UTOF ;



Type
  TOF_PGEXPORTRTF = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnClose                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TobMaj : Tob;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    QMul : THQuery;
    procedure GrilleDblClick(Sender : TObject);
    procedure NouveauFichier(Sender : TObject);
    procedure DupliquerStd(Sender : TObject);
    procedure DupliquerDos(Sender : TObject);
    procedure DupliquerFichier(Predef : String);
  end ;

Implementation

procedure TOF_PGEXPORTRTF.OnLoad;
begin
  Inherited ;
end;

procedure TOF_PGEXPORTRTF.OnClose;
var
CodeRetour, FileAttrs, i, Rep : Integer;
Crit1, Crit2, DAv, DAp, Defaut, Extension, FileP, Langue, Libelle : string;
NoDossier, NomF, Predef : String;
sr : TSearchRec;
Q : TQuery;
TobMajDossier : Tob;
begin
Inherited;
For i:= 0 to TobMaj.Detail.Count-1 do
    begin
    FileP:= TobMaj.Detail[i].GetValue ('CHEMINFICHIER');
    NomF:= TobMaj.Detail[i].GetValue ('FICHIER');
    Predef:= TobMaj.Detail[i].GetValue ('PREDEFINI');
    If (Predef='DOS') then
       NoDossier:= V_PGI.NoDossier
    else
       NoDossier:= '000000';
    Langue:= TobMaj.Detail[i].GetValue ('LANGUE');
    DAv:= TobMaj.Detail[i].GetValue ('DATEFICHIER');
    FileAttrs:= faAnyFile;
    if (FindFirst (FileP, FileAttrs, sr)=0) then
       DAp:= FormatDateTime ('yyyy-mm-dd hh:nn:ss', FileDateToDateTime (Sr.Time));
    If (DAv<>DAp) and not ((Predef='CEG') and (not DroitCegPaie)) then
       begin
       Rep:= PGIAsk ('Le fichier '+NomF+' a été modifié, voulez vous le mettre'+
                     ' à jour dans la base', Ecran.Caption);
       If (Rep=MrYes) then
          begin
          Q:= OpenSQL ('SELECT *'+
                       ' FROM YFILESTD WHERE'+
                       ' YFS_NOM="'+NomF+'" AND'+
                       ' YFS_PREDEFINI="'+Predef+'" AND'+
                       ' YFS_LANGUE="'+Langue+'"',True);
          Crit1:= Q.FindField ('YFS_CRIT1').AsString;
          Crit2:= Q.FindField ('YFS_CRIT2').AsString;
          Defaut:= Q.FindField ('YFS_BCRIT1').AsString;
          Libelle:= Q.FindField ('YFS_LIBELLE').AsString;
          Extension:= Q.FindField ('YFS_EXTFILE').AsString;
          Ferme (Q);
          CodeRetour:=  AGL_YFILESTD_IMPORT (FileP, 'PAIE', NomF, Extension,
                                             Crit1, Crit2, '', '', '', Defaut,
                                             '-', '-', '-', '-', Langue, Predef,
                                             Libelle);
//Gestion du N° de dossier
          TobMajDossier:= Tob.Create ('YFILESTD', Nil, -1);
          Q:= OpenSQl ('SELECT *'+
                       ' FROM YFILESTD WHERE'+
                       ' YFS_LANGUE="'+Langue+'" AND'+
                       ' YFS_CODEPRODUIT="PAIE" AND'+
                       ' YFS_NOM="'+NomF+'" AND'+
                       ' YFS_NODOSSIER="" AND'+
                       ' YFS_CRIT1="'+Crit1+'" AND'+
                       ' YFS_PREDEFINI="'+Predef+'"',True);
          TobMajDossier.LoadDetailDB ('YFILESTD', '', '', Q, False);
          Ferme (Q);
          If (TobMajDossier.Detail.Count=1) then
             TobMajDossier.Detail[0].PutValue ('YFS_NODOSSIER', NoDossier);
          TobMajDossier.InsertOrUpdateDB (False);
          TobMajDossier.Free;
          ExecuteSQL ('DELETE FROM YFILESTD WHERE'+
                      ' YFS_LANGUE="'+Langue+'" AND'+
                      ' YFS_CODEPRODUIT="PAIE" AND'+
                      ' YFS_NOM="'+NomF+'" AND'+
                      ' YFS_NODOSSIER="" AND'+
                      ' YFS_CRIT1="'+Crit1+'" AND'+
                      ' YFS_PREDEFINI="'+Predef+'"');
          end;
       end;
//suppression des fichiers exportés sur le disque local
    If (FileExists (Pchar (FileP))) then
{PT4
       DeleteFile (pchar (FileP));
}
       begin
       Rep:= PGIAsk ('Voulez-vous supprimer définitivement le fichier ?',
                     Ecran.Caption);
       If (Rep=MrYes) then
          DeleteFile (pchar (FileP));
//FIN PT4
       end;
    end;
TobMaj.Free;
end;

procedure TOF_PGEXPORTRTF.OnArgument (S : String ) ;
var
BInsert : TToolBarButton97;
MenuPop : TPopUpMenu;
begin
Inherited;
{$IFNDEF EAGLCLIENT}
Grille:= THDBGrid (GetControl ('Fliste'));
{$ELSE}
Grille:= THGrid (GetControl ('Fliste'));
{$ENDIF}
if (Grille<>NIL) then
   Grille.OnDblClick:= GrilleDblClick;
QMul:= THQuery (Ecran.FindComponent('Q'));
TobMaj:= Tob.Create ('lesfichiers', Nil, -1);
TobMaj.AddChampSup ('FICHIER', False);
TobMaj.AddChampSup ('CHEMINFICHIER', False);
TobMaj.AddChampSup ('DATEFICHIER', False);
BInsert:= TToolbarButton97 (GetControl ('BInsert'));
If (BInsert<>Nil) then
   BInsert.OnClick:= NouveauFichier;
SetControlVisible ('YFS_CRIT3', False);
SetControlVisible ('TYFS_CRIT3', False);
MenuPop:= TPopUpMenu (GetControl ('MENUDUPLIQUE'));
MenuPop.Items[0].OnClick:= DupliquersTD;
MenuPop.Items[1].OnClick:= DupliquerDos;
{PT2
If (V_PGI.ModePcl='1') then
   SetControlText ('XX_WHERE', 'NOT (YFS_PREDEFINI="DOS" AND'+
                               ' YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")');
SetControlText ('XX_WHERE', 'NOT (YFS_CRIT1 = "DUC")'); // PT1
}
{PT3
If (V_PGI.ModePcl='1') then
   SetControlText ('XX_WHERE', 'NOT (YFS_PREDEFINI="DOS" AND'+
                               ' YFS_NODOSSIER<>"'+V_PGI.NoDossier+'") AND'+
                               ' NOT (YFS_CRIT1 = "DUC")')
else
   SetControlText ('XX_WHERE', 'NOT (YFS_CRIT1 = "DUC")'); // PT1
}
If (V_PGI.ModePcl='1') then
   SetControlText ('XX_WHERE', 'NOT (YFS_PREDEFINI="DOS" AND'+
                               ' YFS_NODOSSIER<>"'+V_PGI.NoDossier+'") AND'+
                               ' ((YFS_CRIT1 <> "DUC") AND'+
                               ' (YFS_CRIT1 <> "DAD"))')
else
   SetControlText ('XX_WHERE', ' ((YFS_CRIT1 <> "DUC") AND'+
                               ' (YFS_CRIT1 <> "DAD"))'); // PT1
//FIN PT3
//FIN PT2
end;

procedure TOF_PGEXPORTRTF.GrilleDblClick(Sender : TObject);
var CodeRetour : Integer;
    FileP,NomF,Crit1,Crit2,Langue,Predef : String;
    sr: TSearchRec;
    FileAttrs: Integer;
    DAv : TDateTime;
    T : Tob;
    Ext : String;
    MonExcel : OleVariant;
    bIsNew_p : Boolean;
begin
    {$IFDEF EAGLCLIENT}
    TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
    {$ENDIF}
    NomF := QMul.FindField('YFS_NOM').AsString;
    Crit1 := QMul.FindField('YFS_CRIT1').AsString;
    Crit2 := QMul.FindField('YFS_CRIT2').AsString;
//    Crit3 := QMul.FindField('YFS_CRIT3').AsString;
    Predef :=  QMul.FindField('YFS_PREDEFINI').AsString;
    Langue := QMul.FindField('YFS_LANGUE').AsString;
    Ext := QMul.FindField('YFS_EXTFILE').AsString;
    CodeRetour := AGL_YFILESTD_EXTRACT (FileP, 'PAIE', NomF , Crit1, Crit2,'','','',False,Langue,Predef);
    FileAttrs := faAnyFile;
    if FindFirst(FileP, FileAttrs, sr) = 0 then
    begin
           DAv := FileDateToDateTime(Sr.Time);
    end;
    If (Predef = 'CEG') and (not DroitCegPaie) then PGIBox('Vous ouvrez un fichier CEGID, aucune modification ne sera enregistrée',Ecran.Caption);
    ShellExecute (0, pchar ('open'), PChar(FileP), nil, nil, SW_RESTORE);
    T := TobMaj.FindFirst(['FICHIER,PREDEFINI,LANGUE'],[NomF,Predef,Langue],False);
    If t <> nil then T.PutValue('DATEFICHIER',FormatDateTime('yyyy-mm-dd hh:nn:ss', DAv))
    else
    begin
         T := Tob.Create('lesfichiers',TobMaj,-1);
         T.AddChampSupValeur('FICHIER',NomF);
         T.AddChampSupValeur('PREDEFINI',Predef);
         T.AddChampSupValeur('LANGUE',Langue);
         T.AddChampSupValeur('CHEMINFICHIER',FileP);
         T.AddChampSupValeur('DATEFICHIER',FormatDateTime('yyyy-mm-dd hh:nn:ss', DAv));
    end;
end;

procedure TOF_PGEXPORTRTF.NouveauFichier;
begin
     AglLanceFiche('PAY', 'IMPORTRTF', '', '', '');
     TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGEXPORTRTF.DupliquerStd(Sender : TObject);
begin
     DupliquerFichier('STD');
end;

procedure TOF_PGEXPORTRTF.DupliquerDos(Sender : TObject);
begin
     DupliquerFichier('DOS');
end;


procedure TOF_PGEXPORTRTF.DupliquerFichier(Predef : String);
var Guid,newGuid : String;
    NoDossier,Crit1,Langue,Nom : String;
    LibErreur : String;
    TobYFile,TobNFile,TobNFilePart,T : Tob;
    Q : TQuery;
    i,Nb : Integer;
begin
     if QMul = nil then Exit;
       if (GRILLE.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
         begin
             PGIBox('Aucun élément sélectionné', Ecran.Caption);
             Exit;
         end;
    If Predef = 'STD' then LibErreur := 'en standard'
    else LibErreur := 'pour ce dossier';
    if ((GRILLE.nbSelected) > 0) and (not GRILLE.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', GRILLE.nbSelected, FALSE, TRUE);
    for Nb := 0 to GRILLE.NbSelected - 1 do
    begin
         {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
         {$ENDIF}
         GRILLE.GotoLeBOOKMARK(Nb);
         {$IFDEF EAGLCLIENT}
         TFmul(Ecran).Q.TQ.Seek(GRILLE.Row - 1);
          {$ENDIF}
          Guid := QMul.FindField('YFS_FILEGUID').AsString;
          Crit1 := QMul.FindField('YFS_CRIT1').AsString;
          Langue := QMul.FindField('YFS_LANGUE').AsString;
          Nom := QMul.FindField('YFS_NOM').AsString;
          If Predef = 'DOS' then NoDossier := V_PGI.NoDossier
          else NoDossier := '000000';
          iF ExisteSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CRIT1="'+Crit1+'" AND YFS_LANGUE="'+Langue+'" AND YFS_NOM="'+Nom+'" AND '+
          'YFS_PREDEFINI="'+Predef+'" AND YFS_NODOSSIER="'+Nodossier+'"') then
          begin
               PGIInfo('Ce fichier existe déja '+LibErreur+', duplication impossible','Duplication');
               Continue;
          end;
          NewGuid := AglGetGuid();

          BeginTrans;
          Try
             begin
                 TobYFile := Tob.Create('YFILESTD',Nil,-1);
                 Q := OpenSQL('SELECT * FROM YFILESTD WHERE YFS_FILEGUID="'+Guid+'"',True);
                 TobYFile.LoadDetailDB('YFILESTD','','',Q,False);
                 Ferme(Q);
                 if TobYFile.Detail.Count > 0 then
                 begin
                      TobYFile.Detail[0].PutValue('YFS_FILEGUID',NewGuid);
                      TobYFile.Detail[0].PutValue('YFS_NUMVERSION',0);
                      TobYFile.Detail[0].PutValue('YFS_PREDEFINI',Predef);
                      TobYFile.Detail[0].PutValue('YFS_NODOSSIER',NoDossier);
                      TobYFile.Detail[0].InsertDB(Nil);
                 end;
                 TobYFile.Free;
                 TobNFile := Tob.Create('NFILES',Nil,-1);
                 Q := OpenSQL('SELECT * FROM NFILES WHERE NFI_FILEGUID="'+Guid+'"',True);
                 TobNFile.LoadDetailDB('NFILES','','',Q,False);
                 Ferme(Q);
                 if TobNFile.Detail.Count > 0 then
                 begin
                      TobNFile.Detail[0].PutValue('NFI_FILEGUID',NewGuid);
                      TobNFile.Detail[0].InsertDB(Nil);
                 end;
                 TobNFile.Free;
                 TobNFilePart := Tob.Create('NFILEPARTS',Nil,-1);
                 Q := OpenSQL('SELECT * FROM NFILEPARTS WHERE NFS_FILEGUID="'+Guid+'"',True);
                 TobNFilePart.LoadDetailDB('NFILEPARTS','','',Q,False);
                 Ferme(Q);
                 For i := 0 to TobNFilePart.Detail.Count - 1 do
                 begin
                      TobNFilePart.Detail[i].PutValue('NFS_FILEGUID',NewGuid);
                      TobNFilePart.Detail[i].SetAllModifie(True);
                      TobNFilePart.Detail[i].InsertOrUpdateDB(False);
                 end;
                  TobNFilePart.Free;
                 CommitTrans;
             end;
          except
                RollBack;
          end;
          MoveCurProgressForm(Nom);
         end;
         FiniMoveProgressForm;
    end
    else if Grille.AllSelected then
    begin
         {$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
         {$ENDIF}
         InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
         QMul.First;
         while not QMul.EOF do
         begin
              Guid := QMul.FindField('YFS_FILEGUID').AsString;
          Crit1 := QMul.FindField('YFS_CRIT1').AsString;
          Langue := QMul.FindField('YFS_LANGUE').AsString;
          Nom := QMul.FindField('YFS_NOM').AsString;
          Nodossier := V_PGI.NoDossier;
          iF ExisteSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CRIT1="'+Crit1+'" AND YFS_LANGUE="'+Langue+'" AND YFS_NOM="'+Nom+'" AND '+
          'YFS_PREDEFINI="'+Predef+'" AND YFS_NODOSSIER="'+Nodossier+'"') then
          begin
               PGIInfo('Ce fichier existe déja '+LibErreur+', duplication impossible','Duplication');
               Continue;
          end;
          NewGuid := AglGetGuid();

          BeginTrans;
          Try
             begin
                 TobYFile := Tob.Create('YFILESTD',Nil,-1);
                 Q := OpenSQL('SELECT * FROM YFILESTD WHERE YFS_FILEGUID="'+Guid+'"',True);
                 TobYFile.LoadDetailDB('YFILESTD','','',Q,False);
                 Ferme(Q);
                 if TobYFile.Detail.Count > 0 then
                 begin
                      TobYFile.Detail[0].PutValue('YFS_FILEGUID',NewGuid);
                      TobYFile.Detail[0].PutValue('YFS_NUMVERSION',0);
                      TobYFile.Detail[0].PutValue('YFS_PREDEFINI',Predef);
                      TobYFile.Detail[0].InsertDB(Nil);
                 end;
                 TobYFile.Free;
                 TobNFile := Tob.Create('NFILES',Nil,-1);
                 Q := OpenSQL('SELECT * FROM NFILES WHERE NFI_FILEGUID="'+Guid+'"',True);
                 TobNFile.LoadDetailDB('NFILES','','',Q,False);
                 Ferme(Q);
                 if TobNFile.Detail.Count > 0 then
                 begin
                      TobNFile.Detail[0].PutValue('NFI_FILEGUID',NewGuid);
                      TobNFile.Detail[0].InsertDB(Nil);
                 end;
                 TobNFile.Free;
                 TobNFilePart := Tob.Create('NFILEPARTS',Nil,-1);
                 Q := OpenSQL('SELECT * FROM NFILEPARTS WHERE NFS_FILEGUID="'+Guid+'"',True);
                 TobNFilePart.LoadDetailDB('NFILEPARTS','','',Q,False);
                 Ferme(Q);
                 For i := 0 to TobYFile.Detail.Count - 1 do
                 begin
                      TobNFilePart.Detail[i].PutValue('NFS_FILEGUID',NewGuid);
                      TobNFilePart.Detail[i].SetAllModifie(True);
                      TobNFilePart.Detail[i].InsertDB(Nil);
                 end;
                  TobNFilePart.Free;
                 CommitTrans;
             end;
          except
                RollBack;
          end;
              MoveCurProgressForm(Nom);
              QMul.Next;
         end;
         FiniMoveProgressForm;
    end;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGEXPORTRTF ] ) ;
end.

