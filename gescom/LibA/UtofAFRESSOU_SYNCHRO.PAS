{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 08/04/2003
Modifié le ... : 08/04/2003
Description .. : Source TOF de la TABLE : AFRESSOU_SYNCHRO ()
Mots clefs ... : TOF;AFRESSOU_SYNCHRO
*****************************************************************}
Unit UTOFAFRESSOU_SYNCHRO ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,  Maineagl,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main ,
{$ENDIF}
      forms, sysutils, GCMZSUtil, Utob, UtilMultrt, Windows, Hstatus, HTB97,
      ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
      Dicobtp, SaisUtil, EntGC, M3FP, ParamSoc,
      UtilRessource, UTofAFRes_SyncWhat,
      utilpgi,AglInit,UtilGc,TraducAffaire,UtofAfTraducChampLibre,FileCtrl, CBPPath;

type
  TOF_AFRESSOU_SYNCHRO = class (TOF_AFTRADUCCHAMPLIBRE) // mcd 05/06/02 pour traduciton ne doit pas descendre de TOF
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    FLogName : String;
    procedure BEffaceFicLogOnClick(Sender: TObject);
    procedure BFicLogOnClick(Sender: TObject);
    procedure GetNomFichLog;
    procedure Log(Msg : String);
    procedure HistorizeLesPrix(T : TOB);
    function GetTauxHoraire(Salarie : String; D1, D2 : TDate; Loaded : boolean) : Double;
    procedure AlimFrom(T : TOB; From, Champ : String);

    procedure BalanceLa5ro;
    procedure SynchraleDesRessources;
  end ;

procedure AFLanceFiche_Synchro_RessSalarie;


Implementation
uses ShellAPI;

procedure TOF_AFRESSOU_SYNCHRO.OnArgument (S : String ) ;
begin
  fMulDeTraitement := True;
  Inherited ;

  TToolBarButton97(GetControl('BFICLOG')).OnClick := BFicLogOnClick;
  TToolBarButton97(GetControl('BEFFACEFICLOG')).OnClick := BEffaceFicLogOnClick;
{$IFDEF BTP}
  THValComboBox(GetControl('ARS_TYPERESSOURCE')).enabled:=False;
{$ENDIF}

end ;

procedure TOF_AFRESSOU_SYNCHRO.OnClose ;
begin
  Inherited ;
end ;


procedure TOF_AFRESSOU_SYNCHRO.BEffaceFicLogOnClick(Sender: TObject);
var stDocWord : string;
begin
   stDocWord := GetControlText('FICHIERLOG');
   if (stDocWord = '') then exit;
   if (PGIAskAF('Confirmez-vous la suppression du fichier '+stDocWord+' ?', '') = mrYes) then
    if not DeleteFile(PChar(stDocWord)) then
      PGIBoxAF('Le fichier ne peut-être supprimé ou n''existe pas', '');
end;

procedure TOF_AFRESSOU_SYNCHRO.BFicLogOnClick(Sender: TObject);
var stDocWord : string;
begin
   stDocWord := GetControlText('FICHIERLOG');
   if (stDocWord = '') then exit;
   if not FileExists(stDocWord)
     then PGIBoxAF('Le fichier '+stDocWord+' n''existe pas', '')
     else ShellExecute (0, 'open', PChar(stDocWord), nil, nil, SW_SHOW);
end;

procedure TOF_AFRESSOU_SYNCHRO.GetNomFichLog;
var s,StChemin : string;
begin
   s := GetControlText('FICHIERLOG');
   if s = '' then
   begin
    StChemin := ExtractFilePath(ParamStr(0));
    if (StChemin = '') or not DirectoryExists(StChemin) then

    //Modif FV passage Version 5 - Version 8
    //StChemin := 'C:\PGI00\';
    StChemin := IncludeTrailingBackSlash(TCBPPath.GetCegidUserLocalAppData);

    if ctxScot in V_PGI.PGIContexte then s := StChemin + 'GIRessourceSync.log'
                                    else s := StChemin + 'GARessourceSync.log';
    SetControlText('FICHIERLOG', s);
   end;
   FLogName := s;
end;

procedure TOF_AFRESSOU_SYNCHRO.Log(Msg : String);
var F : TextFile;
begin
   AssignFile(F, FLogName);
   if FileExists(FLogName) then Append(F)
                           else ReWrite(F);
   Writeln(F, Msg);
   CloseFile(F);
end;

procedure TOF_AFRESSOU_SYNCHRO.HistorizeLesPrix(T : TOB);
const Fieldz : Array[1..5] of string = ('ARS_DATEPRIX', 'ARS_TAUXREVIENTUN', 'ARS_PVHT', 'ARS_PVTTC', 'ARS_PVHTCALCUL');
var i, j : integer;
    s : string;
begin
   for i := 4 downto 2 do
    for j := Low(Fieldz) to High(Fieldz) do
    begin
       if i = 2 then s := '' else s := inttostr(i-1);
       T.PutValue(Fieldz[j]+inttostr(i), T.GetValue(Fieldz[j]+s));
    end;
end;

function TOF_AFRESSOU_SYNCHRO.GetTauxHoraire(Salarie : String; D1, D2 : TDate; Loaded : boolean) : Double;
var Q : TQuery;
    Brut, Charges, NbHeures : Double;
begin
   Brut := 0; Charges := 0; NbHeures := 0;
   Q := OpenSQL('SELECT SUM(PPU_CBRUT) AS BRUT, SUM(PPU_CHEURESTRAV) AS HEURES, SUM(PPU_CCOUTPATRON) AS CHARGES FROM PAIEENCOURS WHERE PPU_SALARIE="'+Salarie+'" AND PPU_DATEFIN >= "'+USDateTime(D1)+'" AND PPU_DATEDEBUT <= "'+USDateTime(D2)+'"', true);
   if not Q.EOF then
   begin
      Brut := Q.FindField('BRUT').AsFloat;
      if Loaded then Charges := Q.FindField('CHARGES').AsFloat;
      NbHeures := Q.FindField('HEURES').AsFloat;
   end;
   Ferme(Q);
   if NbHeures = 0
     then result := 0
     else result := (Brut + Charges) / NbHeures;
end;

procedure TOF_AFRESSOU_SYNCHRO.AlimFrom(T : TOB; From, Champ : String);
var s : string;
    Q : TQuery;
    d1, d2 : TDate;
    cp : boolean;
begin
   if Copy(From, 1, 3) = 'THO' then
   begin
      Q := OpenSQL('SELECT PSA_TAUXHORAIRE FROM SALARIES WHERE PSA_SALARIE="'+T.GetValue('ARS_SALARIE')+'"', true);
      T.PutValue(Champ, Q.FindField('PSA_TAUXHORAIRE').AsFloat);
      Ferme(Q);
   end else
   begin
      s := From;
      ReadTokenPipe(s, ',');
      d1 := StrToDate(ReadTokenPipe(s, ','));
      d2 := StrToDate(ReadTokenPipe(s, ','));
      cp := (s = 'X');

      T.PutValue(Champ, GetTauxHoraire(T.GetValue('ARS_SALARIE'), d1, d2, cp));
      (*if Copy(From, 1, 3) = 'THI' then
      begin
         T.PutValue(Champ, GetTauxHoraire(T.GetValue('ARS_SALARIE'), d1, d2, cp, false));
      end else
      if Copy(From, 1, 3) = 'TMO' then
      begin
         T.PutValue(Champ, GetTauxHoraire(T.GetValue('ARS_SALARIE'), d1, d2, cp, true));
      end;*)
   end;
end;

procedure TOF_AFRESSOU_SYNCHRO.BalanceLa5ro;
var TobRes, TobDet : TOB;
    Paramz, SQL, s : string;
    SL : TStringList;
    i, j, nt, anomalies : integer;
    Q : TQuery;
begin
   GetNomFichLog;

   Paramz := AGLLanceFiche('AFF','AFRES_SYNCWHAT', '', '', '');
   if Paramz = '' then exit;

   Log('*** Alignement Ressources / Salariés - Traitement du '+DateTimeToStr(Now));
   //Log('Critères d''alignement : '+Paramz);

   TobRes := TOB.Create('Mom_RESSOURCE', nil, -1);

   SQL := 'SELECT * FROM RESSOURCE'; // A cause du SetAllModifie
   if TraiteEnregMulTable(TFMul(Ecran), SQL ,'ARS_RESSOURCE', 'RESSOURCE', 'ARS_RESSOURCE', 'RESSOURCE', TobRes, true) <> tteOK
     then begin Log('! Pas de sélection !');
                TobRes.Free;
                exit; end;

   SL := TStringList.Create;
   SL.Text := StringReplace(Paramz, ';', Chr(VK_RETURN), [rfReplaceAll]);
   if SL.Values['COMMONFIELDS'] = 'X' then // Alignement des champs communs
     Log('Alignement des champs communs (Nom, Prénom...) de la fiche salarié');
   if SL.Values['ALIMCOST'] = 'X' then
   begin
     if Copy(SL.Values['COSTFROM'], 1, 3) = 'THO' then
       Log('Cout unitaire alimenté par taux horaire');
     if Copy(SL.Values['COSTFROM'], 1, 3) = 'THI' then
       Log('Cout unitaire alimenté par Taux historisé du dernier bulletin');
     if Copy(SL.Values['COSTFROM'], 1, 3) = 'TMO' then
       Log('Cout unitaire alimenté par Taux moyen de la période');
   end;
   if SL.Values['ALIMRATE'] = 'X' then
   begin
     if Copy(SL.Values['RATEFROM'], 1, 3) = 'THO' then
       Log('Prix de revient alimenté par taux horaire');
     if Copy(SL.Values['RATEFROM'], 1, 3) = 'THI' then
       Log('Prix de revient alimenté par Taux historisé du dernier bulletin');
     if Copy(SL.Values['RATEFROM'], 1, 3) = 'TMO' then
       Log('Prix de revient alimenté par Taux moyen de la période');
   end;
   if Copy(SL.Values['HISTO'], 1, 1) = 'X' then
     Log('Historisation des prix');

   nt := 0;
   anomalies := 0;
   InitMove(TobRes.Detail.Count, '');
   for i := 0 to TobRes.Detail.Count-1 do
   begin
      TobDet := TobRes.Detail[i];

      if (TobDet.GetValue('ARS_SALARIE') = null) or (TobDet.GetValue('ARS_SALARIE') = '') then
      begin
         Log('La ressource '+TobDet.GetValue('ARS_RESSOURCE')+' n''a pas de salarié associé');
         inc(anomalies);
      end else
      if not ExisteSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+TobDet.GetValue('ARS_SALARIE')+'"') then
      begin
         Log('Le salarié '+TobDet.GetValue('ARS_SALARIE')+' associé à la ressouce '+TobDet.GetValue('ARS_RESSOURCE')+' n''existe pas');
         inc(anomalies);
      end else
      begin
         if SL.Values['COMMONFIELDS'] = 'X' then // Alignement des champs communs
         begin
            SQL := '';
            for j := Low(RSLinkedFields) to High(RSLinkedFields) do
            begin if SQL <> '' then SQL := SQL+', ';
                  SQL := SQL+RSLinkedFields[j, 1]; end;

            Q := OpenSQL('SELECT '+SQL+' FROM SALARIES WHERE PSA_SALARIE="'+TobDet.GetValue('ARS_SALARIE')+'"', true);
            if not Q.EOF then
              for j := Low(RSLinkedFields) to High(RSLinkedFields) do
                TobDet.PutValue(RSLinkedFields[j, 0], Q.FindField(RSLinkedFields[j, 1]).AsString);
            Ferme(Q);
         end;

         if (Copy(SL.Values['HISTO'], 1, 1) = 'X') {and ((SL.Values['ALIMCOST'] = 'X') or (SL.Values['ALIMRATE'] = 'X'))} then // Gestion de l'historik
         begin
            if StrToDate(Copy(SL.Values['HISTO'], 3, 10)) <= TobDet.GetValue('ARS_DATEPRIX') then
            begin
               Log('Ressource '+TobDet.GetValue('ARS_RESSOURCE')+' : vous avez demandé une historisation des prix au '+Copy(SL.Values['HISTO'], 3, 10)+', alors que la précédente date du '+DateToStr(TobDet.GetValue('ARS_DATEPRIX')));
               inc(anomalies);
               Continue;
            end else
            begin
               HistorizeLesPrix(TobDet);
               TobDet.PutValue('ARS_DATEPRIX', StrToDate(Copy(SL.Values['HISTO'], 3, 10)));
            end;
         end;

         if SL.Values['ALIMCOST'] = 'X' then // Alignement du cout unitaire (TAUXUNIT)
         begin
            AlimFrom(TobDet, SL.Values['COSTFROM'], 'ARS_TAUXUNIT');
         end;

         if TobDet.GetValue('ARS_CALCULPR') = 'X' then // Recalcul du PR
         begin
            TobDet.PutValue('ARS_TAUXREVIENTUN',
                Arrondi(TobDet.GetValue('ARS_TAUXUNIT') *
                        TobDet.GetValue('ARS_TAUXFRAISGEN1') *
                        TobDet.GetValue('ARS_TAUXFRAISGEN2') *
                        TobDet.GetValue('ARS_TAUXCHARGEPAT') *
                        TobDet.GetValue('ARS_COEFMETIER'),
                GetParamSoc('SO_DECPRIX')));
         end else                                       // ou
         if SL.Values['ALIMRATE'] = 'X' then // Alignement du prix de revient (TAUXREVIENTUN)
         begin
            AlimFrom(TobDet, SL.Values['RATEFROM'], 'ARS_TAUXREVIENTUN');
         end;

         if TobDet.GetValue('ARS_CALCULPV') = 'X' then // Recalcul PV
         begin
            TobDet.PutValue('ARS_PVHTCALCUL',
                Arrondi(TobDet.GetValue('ARS_COEFPRPV') *
                        TobDet.GetValue('ARS_TAUXREVIENTUN'),
                GetParamSoc('SO_DECPRIX')));
         end;

         TobDet.PutValue('ARS_DATEMODIF', V_PGI.DateEntree);
         TobDet.SetAllModifie(true);
         TobDet.UpdateDB(False, False);
         inc(nt);
      end;
      MoveCur(False);
   end;
   FiniMove;

   s := 'Fin du traitement - '+inttostr(nt)+' enregistrement(s) aligné(s) sur '+inttostr(TobRes.Detail.Count);
   Log(s);
   Log('');
   if anomalies > 0 then s:= s+Chr(VK_RETURN)+inttostr(anomalies)+' anomalie(s) - consultez le fichier '+FLogName;
   PGIInfoAf(s, '');

   SL.Free;
   TobRes.Free;
end;


procedure TOF_AFRESSOU_SYNCHRO.SynchraleDesRessources;
var F : TFMul;
begin
   F := TFMul(Ecran);
   if (F.FListe.NbSelected=0) and (not F.FListe.AllSelected)
     then MessageAlerte('Aucun élément sélectionné')
     else BalanceLa5ro;
end;

/////////////// Procedure appeleé par le bouton Validation //////////////
procedure AGLSynchroDesRessources(Parms : array of variant; nb : integer);
var  F : TForm;
     TOTOF : TOF;
begin
   F := TForm(longint(Parms[0]));
   if (F is TFMul) then TOTOF := TFMul(F).LaTOF else exit;
   if (TOTOF is TOF_AFRESSOU_SYNCHRO) then TOF_AFRESSOU_SYNCHRO(TOTOF).SynchraleDesRessources else exit;
end;


procedure AFLanceFiche_Synchro_RessSalarie;
begin
   AGLLanceFiche('AFF','AFRESSOU_SYNCHRO','','','');
end;

initialization
  RegisterClasses ([TOF_AFRESSOU_SYNCHRO]);
  RegisterAglProc('SynchroDesRessources', true, 0, AGLSynchroDesRessources);
end.
