{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 10/12/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : Unité commune entre le DP et JURI pour la gestion des évènements
Mots clefs ... : DP;JURI
*****************************************************************}
unit DpJurOutilsEve;

interface

uses
   HCtrls, HQry,
   HEnt1, HStatus, HMsgBox, SysUtils, Controls,
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   Mul,
{$ENDIF}
   HDB, UTob;


/////////// DECLARATION ////////////

// Type "Evènement de suppression" pour gérer suppressions dans DP/Juri
Type
  TEvtSup = Class
    sCode_c     : string;
    sPrefixe_c  : string;
    sTable_c    : string;
    sTypeDos_c  : string;
    sFonction_c : string;
    sGuidPer_c  : string;
    bConfirm_c  : boolean;
  public
    procedure SupprEnreg();
    procedure SuppressionEnregEve;
    procedure SupprimeDocEvt;
  end;

/////////// ENTETES DE FONCTIONS ////////////
procedure CalculerInterval(var dtDate_p : TDateTime;
                           sTypeInterval_p : string;
                           iInterval_p : integer);

function  SupprimeEnreg(sCode_p, sTypeDos_p, sFonction_p, sGuidPer_p, sPrefixe_p,
                        sTable_p : string; bConfirm_p : boolean = false) : boolean;
procedure SupprimeListeEnreg (frmMul : TFMul; St:string; bDoc_p : boolean = false);

/////////// IMPLEMENTATION ////////////

implementation

uses
   UtilMulTraitmt, DpJurOutils;


{Type TEvtSup}

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : SupprEnreg
Description .. : supression d'un enregistrement
Paramètres ... :
*****************************************************************}
procedure TEvtSup.SupprEnreg();
var
    res : integer;
    sWhere: String;
begin
   sWhere := '';
   // tables avec clé composée contenant _GUIDPER + _NOORDRE
   if (sTable_c='DPMVTCAP') or (sTable_c='DPCONTROLE') then
     begin
     sWhere := sTable_c + ' WHERE ' + sPrefixe_c + '_GUIDPER = "' + sGuidPer_c + '" AND ' +
                        sPrefixe_c + '_NOORDRE= "'+ sCode_c + '"' ;
     res := ExecuteSQL ('DELETE FROM '+sWhere);
     if res=-1 then V_PGI.IoError:=oeUnknown ;
     end
   else if (sTable_c='HISTOTITRES') then
     begin
     sWhere := sTable_c +
               ' WHERE ' + sPrefixe_c + '_GUIDPERDOS = "' + sCode_c + '" AND ' +
               sPrefixe_c + '_TITGUIDPER = "' + sGuidPer_c + '" AND ' +
               sPrefixe_c + '_TITNOCPT   = ' + sTypeDos_c + '   AND ' +
               sPrefixe_c + '_NOORDRE    = '+ sFonction_c;
     res := ExecuteSQL ('DELETE FROM ' + sWhere);
     if res=-1 then V_PGI.IoError:=oeUnknown ;
     end
   // table ANNULIEN
   else if (sTable_c = 'ANNULIEN') then
     begin
     sWhere := sTable_c + ' WHERE ANL_GUIDPERDOS = "' + sCode_c + '"'
                        + ' AND ANL_TYPEDOS = "' + sTypeDos_c + '"'
                        + ' AND ANL_NOORDRE = 1 AND ANL_FONCTION = "' + sFonction_c + '"'
                        + ' AND ANL_GUIDPER = "'+ sGuidPer_c + '"';
     res := ExecuteSQL ('DELETE FROM ' + sWhere);
     if res=-1 then V_PGI.IoError:=oeUnknown ;
     end
   // table JUEVENEMENT
   else if (sTable_c = 'JUEVENEMENT') then
     begin
     sWhere := sTable_c + ' where JEV_GUIDEVT = "' + sCode_c + '"';
     SuppressionEnregEve;
     end
   // tables avec _GUIDPER
   else
     begin
     sWhere := sTable_c + ' where ' + sPrefixe_c + '_GUIDPER = "' + sCode_c + '"';
     res := ExecuteSQL ('DELETE FROM '+sWhere);
     if res=-1 then V_PGI.IoError:=oeUnknown ;
     end;
end;

{*****************************************************************
Auteur ....... : ??
Date ......... : ??/??/??
Procédure .... : SuppressionEnregEve
Description .. :
Paramètres ... :
*****************************************************************}
procedure TEvtSup.SuppressionEnregEve;
begin
   beginTrans;
   try
      SupprimeDocEvt;
    // #### rajouter suppression des mails associés (si non attachés à la GED)
      if V_PGI.IoError=oeOK then
         ExecuteSQL ('DELETE FROM JUEVENEMENT WHERE JEV_GUIDEVT = "' + sCode_c + '"');
   except
      V_PGI.IoError:=oeUnknown;
   end;
   if V_PGI.IoError=oeUnknown then
      RollBack
   else
      CommitTrans;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : ??/??/??
Procédure .... : SupprimeDocEvt
Description .. :
Paramètres ... :
*****************************************************************}
procedure TEvtSup.SupprimeDocEvt;
Var
   QQ : TQuery ;
   sNomDoc{,sTexte} : string;
   {bRet, bDelete : boolean;}
begin
   if sCode_c = '' then exit;
   QQ:=OpenSQL('SELECT JEV_FAMEVT,JEV_DOCNOM FROM JUEVENEMENT WHERE JEV_GUIDEVT = "' + sCode_c + '"',TRUE) ;
   {bRet:=true;}
   if Not QQ.EOF then
   begin
      QQ.First;
      sNomDoc := QQ.FindField('JEV_DOCNOM').AsString;
      if (QQ.FindField('JEV_FAMEVT').AsString='DOC') and (sNomDoc<>'') then
      begin
         // MD 02/05/03 - pas de messages dans une transaction => choix fait avant !
         if bConfirm_c then
{         begin
            sTexte := 'Voulez-vous détruire également le fichier document "'+sNomDoc+'" attaché à cet événement ?';
            bDelete := HShowMessage('0;Suppression des événements;'+sTexte+';Q;YN;N;N;','','')=mrYes;
         end
         else
            bDelete := not bConfirm_c;
         if bDelete then
            bRet:=}DeleteFile(sNomDoc);
      end;
   end;
{  // MD 02/05/03 - pas de messages dans une transaction !
   if not bRet then
     begin
     MessageAlerte('Suppression du fichier '+sNomDoc+' impossible (ou fichier inexistant)') ;
     end; }
   Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CalculerInterval(var dtDate_p : TDateTime;
                           sTypeInterval_p : string;
                           iInterval_p : integer);
var
   iReliqJour_l : integer;
begin
   if iInterval_p = 0 then
      exit;

   if (sTypeInterval_p = 'SEC') or
      (sTypeInterval_p = 'MIN') or
      (sTypeInterval_p = 'HEU') then
   begin
      dtDate_p := PlusHeure(dtDate_p, iInterval_p, Copy(sTypeInterval_p,0,1), iReliqJour_l);
   end
   else if (sTypeInterval_p = 'JOU') or
      (sTypeInterval_p = 'MOI') or
      (sTypeInterval_p = 'ANN') then
      dtDate_p := PlusDate(dtDate_p, iInterval_p, Copy(sTypeInterval_p,0,1) );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : ??/??/??
Procédure .... : SupprimeEnreg
Description .. : transaction
Paramètres ... :
*****************************************************************}
function  SupprimeEnreg(sCode_p, sTypeDos_p, sFonction_p, sGuidPer_p, sPrefixe_p,
                        sTable_p : string; bConfirm_p : boolean = false) : boolean;
var
    EvtSup : TEvtSup;
begin
  Result := false;
  EvtSup := TEvtSup.Create;
  EvtSup.sCode_c := sCode_p;
  EvtSup.sPrefixe_c := sPrefixe_p;
  EvtSup.sTable_c := sTable_p;
  EvtSup.sTypeDos_c := sTypeDos_p;
  EvtSup.sFonction_c := sFonction_p;
  EvtSup.sGuidPer_c := sGuidPer_p;
  EvtSup.bConfirm_c := bConfirm_p;
  if Transactions(EvtSup.SupprEnreg, 3)<>oeOk then
    begin
    PGIInfo('Suppression impossible', TitreHalley) ;
    EvtSup.Free;
    exit;
    end ;
  EvtSup.Free;
  Result := true;
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : SupprimeListeEnreg
Description .. : suppression d'une liste de liens (capital, contrôles fiscaux...)
Paramètres ... : Grille, Requête, Table
*****************************************************************}
procedure SupprimeListeEnreg (frmMul : TFMul; St:string; bDoc_p : boolean = false);
var i : integer;
    Texte : string;
    bConfirme : boolean;
    L : THDBGrid;
    Q : THQuery;
begin
  L := frmMul.FListe;
  Q := frmMul.Q;
  if (L.NbSelected=0) and (Not L.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné', TitreHalley);
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les informations.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression historique;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;

  if bDoc_p then
  begin
     Texte := 'Voulez-vous également détruire les fichiers documents de toutes les lignes sélectionnées ?';
     bConfirme := HShowMessage('0;Suppression des événements;'+Texte+';Q;YN;N;N;','','')=mrYes;
  end;
  
  if L.AllSelected then
    BEGIN
{$IFDEF EAGLCLIENT}
    if not frmMul.FetchLesTous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      Q.First;
      while Not Q.EOF do
        begin
        MoveCur(False);
        if St ='DPMVTCAP' then
          begin
          if not SupprimeEnreg(Q.FindField('DPM_NOORDRE').AsString, '', '', Q.FindField('DPM_GUIDPER').AsString, 'DPM', 'DPMVTCAP') then Break ;
          end;
        if St ='HISTOTITRES' then
          begin
          if not SupprimeEnreg(Q.FindField('TIT_GUIDPERDOS').AsString,
                               Q.FindField('TIT_TITNOCPT').AsString,
                               Q.FindField('TIT_NOORDRE').AsString,
                               Q.FindField('TIT_TITGUIDPER').AsString, 'TIT', 'HISTOTITRES') then Break ;
          end;
        if St ='DPCONTROLE' then
          begin
          if not SupprimeEnreg(Q.FindField('DCL_NOORDRE').AsString, '', '', Q.FindField('DCL_GUIDPER').AsString, 'DCL', 'DPCONTROLE') then Break ;
          end;
        if St ='ANNULIEN' then
          begin
          if not SupprimeEnreg(Q.FindField('ANL_GUIDPERDOS').AsString, Q.FindField('ANL_TYPEDOS').AsString, Q.FindField('ANL_FONCTION').AsString, Q.FindField('ANL_GUIDPER').AsString, 'ANL', 'ANNULIEN') then Break ;
          end;
        if St = 'JUEVENEMENT' then
          begin
          if not SupprimeEnreg(Q.FindField('JEV_GUIDEVT').AsString, '', '', '',
                             'JEV', 'JUEVENEMENT', bConfirme ) then Break ;
          end;
        // #### en fait, géré par AGLSupprimeListAnnu
        if St ='ANNUAIRE' then
          begin
          if not SupprimeEnreg(Q.FindField('ANN_GUIDPER').AsString, '', '', '', 'ANN', 'ANNUAIRE') then Break ;
          end;
        Q.Next;
        end;
      end;
    END
  ELSE
    BEGIN
    InitMove(L.NbSelected,'');
    for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(L.Row - 1) ;
{$ENDIF}
      if St ='DPMVTCAP' then
        begin
        if not SupprimeEnreg(Q.FindField('DPM_NOORDRE').AsString, '', '', Q.FindField('DPM_GUIDPER').AsString, 'DPM', 'DPMVTCAP') then Break ;
        end;
      if St ='HISTOTITRES' then
        begin
        if not SupprimeEnreg(Q.FindField('TIT_GUIDPERDOS').AsString,
                             Q.FindField('TIT_TITNOCPT').AsString,
                             Q.FindField('TIT_NOORDRE').AsString,
                             Q.FindField('TIT_TITGUIDPER').AsString, 'TIT', 'HISTOTITRES') then Break ;
        end;
      if St ='DPCONTROLE' then
        begin
        if not SupprimeEnreg(Q.FindField('DCL_NOORDRE').AsString, '', '', Q.FindField('DCL_GUIDPER').AsString, 'DCL', 'DPCONTROLE') then Break ;
        end;
      if St ='ANNULIEN' then
        begin
        if not SupprimeEnreg(Q.FindField('ANL_GUIDPERDOS').AsString, Q.FindField('ANL_TYPEDOS').AsString, Q.FindField('ANL_FONCTION').AsString, Q.FindField('ANL_GUIDPER').AsString, 'ANL', 'ANNULIEN') then Break ;
        end;
      if St = 'JUEVENEMENT' then
        begin
        if not SupprimeEnreg(Q.FindField('JEV_GUIDEVT').AsString, '', '', '',
                        'JEV', 'JUEVENEMENT', bConfirme) then Break ;
        end;
      if St ='ANNUAIRE' then
        begin
        if not SupprimeEnreg(Q.FindField('ANN_GUIDPER').AsString, '', '', '', 'ANN', 'ANNUAIRE') then Break ;
        end;
      end;
    FiniMove;
    END;
  // déselectionne
  FinTraitmtMul(frmMul);
end;


end.
