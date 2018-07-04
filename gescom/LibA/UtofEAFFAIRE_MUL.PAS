{***********UNITE*************************************************
Auteur  ...... :  Patrice ARANEGA
Créé le ...... : 12/10/2001
Modifié le ... : 12/10/2001
Description .. : Source TOF de la TABLE : EAFFAIRE_MUL ()
Suite ........ : Intégration d'E- Affaires
Mots clefs ... : TOF;EAFFAIRE_MUL
*****************************************************************}
Unit UtofEAFFAIRE_MUL;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,FileCtrl,
{$IFDEF EAGLCLIENT}
    eMul,Main_eAGL,
{$ELSE}
    db,Mul, Shellapi,Windows,Fe_Main,dbTables,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF,UTOFAfBaseCodeAffaire, UtilMulTrt,UtilEAffaire,
     Utob, M3FP,DicoAf, HTB97, HeureUtil,UTOBXLS ;

Type
  TOF_EAFFAIRE_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    NumErr : integer;
    FichierLog : TextFile;
    TobDetAff, TOBImport : TOB;
    CCLog  : THEdit;
    procedure IntegrationEAffaire ;
    procedure IntegrationEAffaireTotal ;
    procedure ValideUneEAffaire;
    function NomFichierLog : string;
    procedure BFicLogOnClick(Sender: TObject);
    procedure BImportOnClick(Sender: TObject);
    procedure BImportETiersOnClick(Sender: TObject);
  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;

  end ;

// libellés des messages - Attention les numéros sont importants reférence des erreurs
Const	MsgErr: array[1..9] of string 	= (
          {1}        'Code affaire déja utilisé',
          {2}        'Code affaire incorrect',
          {3}        'Code Etat non valide',
          {4}        'Le client est inexistant',
          {5}        'Le client a un état comptable rouge',
          {6}        'Le Client est fermé',
          {7}        'Le client n''est pas renseigné',
          {8}        'Code Devise incorrecte',
          {9}        'Validation de l''affaire impossible'
          );

Implementation

procedure TOF_EAFFAIRE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EAFFAIRE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EAFFAIRE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_EAFFAIRE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EAFFAIRE_MUL.OnArgument (S : String ) ;
Var Bt : TToolBarButton97;
begin
  Inherited ;
CCLog := THEdit(GetControl('FICHIERLOG'));
NomFichierLog;
Bt := TToolBarButton97(GetControl('BFICLOG'));
If Bt <> Nil then Bt.OnClick:= BFicLogOnClick;
Bt := TToolBarButton97(GetControl('BIMPOREXTERNE'));
If Bt <> Nil then Bt.OnClick:= BImportOnClick;
Bt := TToolBarButton97(GetControl('BIMPORTETIERS'));
If Bt <> Nil then Bt.OnClick:= BImportETiersOnClick;
end ;

procedure TOF_EAFFAIRE_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_EAFFAIRE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff1:=THEdit(GetControl('EAF_AFFAIRE1')); Aff2:=THEdit(GetControl('EAF_AFFAIRE2'));
Aff3:=THEdit(GetControl('EAF_AFFAIRE3')); Aff4:=THEdit(GetControl('EAF_AVENANT'));
Tiers:=THEdit(GetControl('EAF_TIERS'));
end;

//******************************************************************************
//*********************** Traitement des E-Affaires ****************************
//******************************************************************************
procedure TOF_EAFFAIRE_MUL.IntegrationEAffaire ;
Var TobEAffs : TOB;
    i, NbAffTot, NbAffValid : integer;
    stErr, NomFicLog  : string;
begin
  TobEAffs := Tob.Create('les E-affaires',NIL,-1);
  try

    // normal de tout prendre
    TraiteEnregMulTable (TFMul(Ecran),'SELECT * FROM EAFFAIRE','EAF_AFFAIRE0;EAF_AFFAIRE1;EAF_AFFAIRE2;EAF_AFFAIRE3;EAF_AVENANT'
                      ,'EAFFAIRE', 'EAF_AFFAIRE1','EAFFAIRE',TobEAffs, True);

    NbAffValid := 0;
    NbAffTot := TobEAffs.Detail.Count;
    if NbAffTot > 0 then
       begin
       NomFicLog:=NomFichierLog; // gestion du fichier log
       AssignFile(FichierLog, NomFicLog);
       if (FileExists(NomFicLog) = TRUE) then  Append (FichierLog) else  Rewrite (FichierLog);
       writeln (FichierLog, '************ Importation des données E-AFFAIRE : '+inttostr(NbAffTot) + ' ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
       end
    else Exit;

    if (TobEAffs.Detail.Count>=0) then TobEAffs.Detail[0].AddChampSup('VALIDE', True);

    For i := 0 To NbAffTot-1 do
       begin
       TobDetAff := TobEAffs.Detail[i];

       Transactions(ValideUneEAffaire,1);

       if V_PGI.IoError <> OeOk then NumErr:=9; // mcd ajout erreur
       if NumErr = 0 then
          begin
          Inc(NbAffValid); TobDetAff.PutValue('VALIDE', 'X');
          end
       else
          begin
          TobDetAff.PutValue('VALIDE', '-');
          stErr := ' Erreur E-Affaire : '+Trim(TobDetAff.GetValue('EAF_AFFAIRE1'))+ ' '+ Trim(TobDetAff.GetValue('EAF_AFFAIRE2'))+ ' '+
                                          Trim(TobDetAff.GetValue('EAF_AFFAIRE3'))+ ' => ' + MsgErr[NumErr];
          Writeln (FichierLog, stErr);
          end;
       end;
    // On supprime dans la table EAFFAIRE les lignes importées dans la table AFFAIRE
    for i := TobEAffs.Detail.Count-1 downto 0 do
       BEGIN
       if (TobEAffs.Detail[i].GetValue('VALIDE')='-')
          then TobEAffs.Detail[i].Free;
       END;
    if Not(TobEAffs.DeleteDB(True)) then PGIBoxAF ('Suppression des lignes E-Affaires impossible','Import E-Affaires');
    // Affichage du message du nombre d'affaires importées
        PGIInfoAf(IntToStr(NbAffValid) + ' affaire(s) intégrée(s) et ' + IntToStr(NbAffTot -NbAffValid) + ' affaires refusée(s)','Import E-Affaires' );

    writeln (FichierLog, '************ Fin de l''importation des données E-Affaire ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
    CloseFile (FichierLog);
  Finally
   if TobEAffs <> nil then TobEAffs.Free;
  end;
end;

// **************** transaction de validation d'une affaire ********************
procedure TOF_EAFFAIRE_MUL.ValideUneEAffaire;
begin
NumErr := TraiteUneEAffaire(TobDetAff);
end;

procedure TOF_EAFFAIRE_MUL.IntegrationEAffaireTotal;
Var TobEAffs,TobEAffsNonOK : TOB;
    i, NbAffTot, NbAffValid , NbAffLot: integer;
    stErr, NomFicLog  : string;
    zret : boolean;
    req : string;
    QQ : Tquery;
begin

	NbAffValid := 0;
  NbAffTot := 0;

//if NbAffTot > 0 then
//   begin
   NomFicLog:=NomFichierLog; // gestion du fichier log
   AssignFile(FichierLog, NomFicLog);
   if (FileExists(NomFicLog) = TRUE) then  Append (FichierLog) else  Rewrite (FichierLog);
   writeln (FichierLog, '************ Importation des données E-AFFAIRE : '+inttostr(NbAffTot) + ' ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
//end
//else Exit;

zret := true;
While zret  do    
BEGIN

  TobEAffs := Tob.Create('les E-affaires',NIL,-1);
  TobEAffsNonOK := Tob.Create('les E-affaires non ok',NIL,-1);
  try
    Req:= 'SELECT top 10 * FROM EAFFAIRE WHERE eaf_boollibre3<>"X"';
    QQ := OpenSQL(Req,True) ;
    if not QQ.EOF then
       begin
        TobEAffs.LoadDetailDB('EAFFAIRE','','',QQ,False,false) ;
       end
    else
        begin
          zret := false;
        end;
    Ferme (QQ);

    if (zret) then
    Begin
      if (TobEAffs.Detail.Count>=0) then TobEAffs.Detail[0].AddChampSup('VALIDE', True);

      NbAfflot := TobEAffs.Detail.Count;
      NbAffTot := NBAffTot + NbAffLot;
      For i := 0 To NbAffLot-1 do
         begin
         TobDetAff := TobEAffs.Detail[i];

         Transactions(ValideUneEAffaire,1);

         if NumErr = 0 then
            begin
            Inc(NbAffValid); TobDetAff.PutValue('VALIDE', 'X');
            end
         else
            begin
            TobDetAff.PutValue('EAF_BOOLLIBRE3', 'X');  // ne pas retriater dans prochioan lot
            TobDetAff.PutValue('VALIDE', '-');

            stErr := ' Erreur E-Affaire : '+Trim(TobDetAff.GetValue('EAF_AFFAIRE1'))+ ' '+ Trim(TobDetAff.GetValue('EAF_AFFAIRE2'))+ ' '+
                                            Trim(TobDetAff.GetValue('EAF_AFFAIRE3'))+ ' => ' + MsgErr[NumErr];
            Writeln (FichierLog, stErr);
            end;
         end;
      // On supprime dans la table EAFFAIRE les lignes importées dans la table AFFAIRE
      for i := TobEAffs.Detail.Count-1 downto 0 do
         BEGIN
         TobDetAff := TobEAffs.Detail[i];
         if (TobDetAff.GetValue('VALIDE')='-') then  TobDetAff.changeParent(TobEAffsNonOk,-1);;
         END;

      if (TobEAffs.Detail.Count <>0) then
        Begin
        if Not(TobEAffs.DeleteDB(True)) then
        begin
          PGIBoxAF ('Suppression des lignes E-Affaires impossible','Import E-Affaires');
          zret := false;
        end;
      end;

      if (TobEAffsNonOK.Detail.Count <>0) then
      Begin
        if Not(TobEAffsNonOK.UpdateDB(True)) then
          begin
            PGIBoxAF ('Maj des lignes E-Affaires impossible','Import E-Affaires');
            zret := false;
          end;
      end;

      END; //zret
    finally
      TobEAffs.Free;
      TobEAffsNonOK.Free;
    end;

END; // while
// Affichage du message du nombre d'affaires importées
PGIInfoAf(IntToStr(NbAffValid) + ' affaire(s) intégrée(s) et ' + IntToStr(NbAffTot -NbAffValid) + ' affaires refusée(s)','Import E-Affaires' );

writeln (FichierLog, '************ Fin de l''importation des données E-Affaire ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
CloseFile (FichierLog);



end;


//******************************************************************************
//************** Visualisation / Traitement du fichier de LOG ******************
//******************************************************************************
procedure TOF_EAFFAIRE_MUL.BFicLogOnClick(Sender: TObject);
var stDocWord : string ;
begin
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
stDocWord:=CCLog.Text;
if (stDocWord='') then exit;
if Not FileExists(stDocWord) then
    begin
    PGIInfoAF('Le fichier ' + stDocWord +' n''existe pas', TFMUL(Ecran).caption);
    exit;
    end;
ShellExecute( 0, PCHAR('open'),PChar(stDocWord), Nil,Nil,SW_RESTORE);
{$ENDIF}
end;

// *************** Recup du nom de fichier de log ******************************
function TOF_EAFFAIRE_MUL.NomFichierLog:string;
var
NomLog,NomFicLog,PathLog : string;
begin
NomLog:=''; NomFicLog:=''; PathLog:=''; Result:='';
NomLog := CCLog.Text;
if (NomLog<>'') then
    begin
    Pathlog:=ExtractFilePath(NomLog) ;
    NomFicLog:=ChangeFileExt(ExtractFileName(NomLog),'.log') ;
    end;
if (PathLog<>'') then
    begin
    if (NomFicLog<>'') then Result :=PathLog+NomFicLog
                       else Result :=PathLog+'ImportEAffaire.log';
    end
else
   Result:='c:\ImportEAffaire.log';
CCLog.Text := Result;
end;

//******************************************************************************
//****************************Import Externe vers E-Affaires *******************
//******************************************************************************
procedure TOF_EAFFAIRE_MUL.BImportOnClick(Sender: TObject);
var
T : TOB;
sNomRep:string;
rep,NomCompletFic:string;
SearchRec: TSearchRec;
begin
NomCompletFic:='';
// pas de tof nom de fct inutile
Rep:= AglLanceFiche('AFF', 'AFEACTIVITEREPIMP','','','EAFFAIRE');
if (Rep='') then Exit;

sNomRep := copy( Rep, 9, length(Rep)-8);
if not DirectoryExists(sNomRep) then begin PGIInfoAf('Le répertoire ' + sNomRep + ' n''existe pas', TFMUL(Ecran).caption); exit; end;


if pos('EXCEL', rep)<>0 then
    begin
    FindFirst(sNomRep+'\*.xls', faAnyFile, SearchRec);
    repeat
    if (SearchRec.Name<>'') then
    if FileExists (sNomRep+'\'+SearchRec.Name) then
        begin
        NomCompletFic := sNomRep+'\'+SearchRec.Name;
        // récupération de la tob
        T := TOB.Create('', Nil, -1);
        try
        ImportTOBFromXLS (T, NomCompletFic, true, true); // titres de la table dans les titres des colonnes
        T.insertDB(nil,false);
        finally
         T.Free;
         end;
        end;
    until FindNext(SearchRec)<>0 ;
    sysutils.FindClose(SearchRec);
    end
else
if pos('ASCII', rep)<>0 then
    begin
    FindFirst(sNomRep+'\*.asc', faAnyFile, SearchRec);
    repeat
    if (SearchRec.Name<>'') then
    if FileExists (sNomRep+'\'+SearchRec.Name) then
        try
        TobImport := TOB.Create('', nil, -1);
        NomCompletFic := sNomRep+'\'+SearchRec.Name;
        if FileExists (NomCompletFic) then
            begin
            TobLoadFromFile( NomCompletFic, Nil, TobImport ) ;
            end;
        finally
            if (TobImport<>nil) then
                if (TobImport.Detail.count <>0) then
                    begin
                    if (TobImport.Detail[0].NomTable<>'EAFFAIRE') then
                        PGIInfoAf('Les lignes importées ne sont pas des lignes de l''e-Affaire. Table : '+TobImport.Detail[0].NomTable, TFMUL(Ecran).caption)
                    else
                        TobImport.InsertDB(nil, true);
                    end;
            TobImport.Free;
        end;
    until FindNext(SearchRec)<>0 ;
    sysutils.FindClose(SearchRec);
    end;

TFMUL(Ecran).ChercheClick;
TToolBarButton97(GetControl('bSelectAll')).Down:=false;
TFMul(Ecran).FListe.AllSelected:=False;
end;

//******************************************************************************
//****************************Import Externe vers E-Affaires *******************
//******************************************************************************
procedure TOF_EAFFAIRE_MUL.BImportETiersOnClick(Sender: TObject);
var
T : TOB;
sNomRep:string;
rep,NomCompletFic:string;
SearchRec: TSearchRec;
begin
NomCompletFic:='';
Rep:= AglLanceFiche('AFF', 'AFEACTIVITEREPIMP','','','ETIERS');
if (Rep='') then Exit;

sNomRep := copy( Rep, 9, length(Rep)-8);
if not DirectoryExists(sNomRep) then begin PGIInfoAf('Le répertoire ' + sNomRep + ' n''existe pas', TFMUL(Ecran).caption); exit; end;


if pos('EXCEL', rep)<>0 then
    begin
    FindFirst(sNomRep+'\*.xls', faAnyFile, SearchRec);
    repeat
    if (SearchRec.Name<>'') then
    if FileExists (sNomRep+'\'+SearchRec.Name) then
        begin
        NomCompletFic := sNomRep+'\'+SearchRec.Name;
        // récupération de la tob
        T := TOB.Create('', Nil, -1);
        try
        ImportTOBFromXLS (T, NomCompletFic, true, true); // titres de la table dans les titres des colonnes
        T.insertDB(nil,false);
        finally
         T.Free;
         end;
        end;
    until FindNext(SearchRec)<>0 ;
    sysutils.FindClose(SearchRec);
    end
else
if pos('ASCII', rep)<>0 then
    begin
    FindFirst(sNomRep+'\*.asc', faAnyFile, SearchRec);
    repeat
    if (SearchRec.Name<>'') then
    if FileExists (sNomRep+'\'+SearchRec.Name) then
        try
        TobImport := TOB.Create('', nil, -1);
        NomCompletFic := sNomRep+'\'+SearchRec.Name;
        if FileExists (NomCompletFic) then
            begin
            TobLoadFromFile( NomCompletFic, Nil, TobImport ) ;
            end;
        finally
            if (TobImport<>nil) then
                if (TobImport.Detail.count <>0) then
                    begin
                    if (TobImport.Detail[0].NomTable<>'ETIERS') then
                        PGIInfoAf('Les lignes importées ne sont pas des lignes de l''e-Tiers. Table : '+TobImport.Detail[0].NomTable, TFMUL(Ecran).caption)
                    else
                        TobImport.InsertDB(nil, true);
                    end;
            TobImport.Free;
        end;
    until FindNext(SearchRec)<>0 ;
    sysutils.FindClose(SearchRec);
    end;

TFMUL(Ecran).ChercheClick;
TToolBarButton97(GetControl('bSelectAll')).Down:=false;
TFMul(Ecran).FListe.AllSelected:=False;
end;

// ************ Appel depuis le script ************************
procedure AGIntegrationEAffaire( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else latof := nil;
if (Latof is TOF_EAFFAIRE_MUL) then TOF_EAFFAIRE_MUL(LaTof).IntegrationEAffaire else exit;
end;
procedure AGIntegrationEAffaireTotal( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else laTof := nil;
if (Latof is TOF_EAFFAIRE_MUL) then TOF_EAFFAIRE_MUL(LaTof).IntegrationEAffaireTotal else exit;
end;

Initialization
  registerclasses ( [ TOF_EAFFAIRE_MUL ] ) ;
  RegisterAglProc( 'IntegrationEAffaire',True,0,AGIntegrationEAffaire);
  RegisterAglProc( 'IntegrationEAffaireTotal',True,0,AGIntegrationEAffaireTotal);
end.
