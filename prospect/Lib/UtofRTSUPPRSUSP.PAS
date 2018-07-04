{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/08/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTSUPPRSUSP ()
Mots clefs ... : TOF;RTSUPPRSUSP
*****************************************************************}
Unit UtofRTSUPPRSUSP ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HTB97,
     ed_tools,
     HMsgBox,
     dialogs,
     utilRT,
     utilGC,
     utobdebug,
     utofcomm,
     UTOF ;

Type
  TOF_RTSUPPRSUSP = Class (tTOFComm)
    public

    btOuvrir : TToolbarButton97;
    cOrigine : THValComboBox;

//FQ10455    Fliste : THGrid;

    procedure FlisteOnDblClick ( Sender : Tobject );
    procedure SupprSusp ( Sender : TObject );
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  end ;

Function RTLanceFiche_RTSUPPRSUSP(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_RTSUPPRSUSP(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTSUPPRSUSP.OnArgument (Arguments : String ) ;
var
  F : TFmul;
begin
  Inherited ;
  F := TFmul(Ecran);
  fTableName := 'SUSPECTS' ;
  fMulDeTraitement := True ;

  btOuvrir := TToolbarButton97(GetControl('BOUVRIR'));
  btOuvrir.OnClick := SupprSusp;

  SetControlChecked('RSU_FERME', True);
  SetControlVisible('BSELECTALL', True);

  //FQ 10455 on enlève la fonction de doubleclick
//  Fliste := THGrid(GetControl('FLISTE'));
//  Fliste.OnDblClick := FlisteOnDblClick;

  SetControlText('XX_WHERE', '(RSU_SUSPECT NOT IN (SELECT RVB_SUSPECT FROM CIBLAGEELEMENT))');

{CRM_CRM10817_CD}
  F.Caption := TraduireMemoire ('Suppression des suspects');
  UpdateCaption(Ecran);
  RTMajChampsLibresSuspect(F);

  //FQ10467 gestion des commerciaux
  if not GereCommercial then
  begin
    SetControlVisible ('RSU_REPRESENTANT', False);
    SetControlVisible ('TRSU_REPRESENTANT', False);
    SetControlVisible ('RSU_ZONECOM', False); //mcd 13972 27/04/07
    SetControlVisible ('TRSU_ZONECOM', False);
  end;

  //FQ 10455
  cOrigine := THValComboBox(GetControl('RSU_ORIGINETIERS'));
  cOrigine.Items.Insert(1,'<<Aucun>>');
  cOrigine.Values.Insert(1,'');





end ;

procedure TOF_RTSUPPRSUSP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTSUPPRSUSP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTSUPPRSUSP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTSUPPRSUSP.OnLoad ;
var
  StrWhere : String;

begin
  Inherited ;
  if GetCheckBoxState('ELEMCIBL') = cbchecked	then
    StrWhere := '(RSU_SUSPECT IN (SELECT RVB_SUSPECT FROM CIBLAGEELEMENT))';

  if GetCheckBoxState('ELEMCIBL') = cbUnchecked then
    StrWhere := '(RSU_SUSPECT NOT IN (SELECT RVB_SUSPECT FROM CIBLAGEELEMENT))';

  if GetCheckBoxState('ELEMCIBL') = cbGrayed then
    StrWhere := '';

  //FQ10455
  if cOrigine.Text = '<<Aucun>>' then
  begin
    if GetCheckBoxState('ELEMCIBL') <> cbGrayed then
      StrWhere := StrWhere + ' AND (RSU_ORIGINETIERS = "")'
    else
      StrWhere := 'RSU_ORIGINETIERS = ""';
  end;
  SetControlText('XX_WHERE', StrWhere);

end ;

procedure TOF_RTSUPPRSUSP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTSUPPRSUSP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTSUPPRSUSP.OnCancel () ;
begin
  Inherited ;
end ;



procedure TOF_RTSUPPRSUSP.SupprSusp(Sender: TObject);
var
   F : TFmul;
   StrWhere1 : String;
   StrWhere2 : String;

begin
  F := Tfmul (ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
       PGIInfo('Aucun élément sélectionné');
       exit;
  end;

  if PGIAsk('Attention !#10#13'+
            'Cette commande va supprimer toutes les fiches sélectionnées.#10#13'+
            'Les fiches complémentaires seront aussi supprimées.#10#13'+
            'Confirmez-vous le traitement ?', 'Suppression suspect') = mrNo then
            EXIT;


  InitMoveProgressForm(nil,'Suppression en cours','',3,True,True);
  if not F.FListe.AllSelected then
  begin
    BEGINTRANS;
    Try
      StrWhere1 := MakeWhere;         // le MakeWhere renvoie le code suspect
      StrWhere2:= StringReplace_(strwhere1,'RSU_','RSC_',[rfReplaceAll]);
      ExecuteSQL('DELETE FROM SUSPECTSCOMPL WHERE '+StrWhere2);
      MoveCurProgressForm('Veuillez patienter');

      StrWhere2:= StringReplace_(strwhere1,'RSU_','RVB_',[rfReplaceAll]);
      ExecuteSQL('DELETE FROM CIBLAGEELEMENT WHERE '+StrWhere2);
      MoveCurProgressForm('Veuillez patienter');

      ExecuteSQL('DELETE FROM SUSPECTS WHERE '+StrWhere1);
      MoveCurProgressForm('Veuillez patienter');
      COMMITTRANS;
    Except
      on E: Exception do ROLLBACK;
    end;
  end
  else
  begin
    BEGINTRANS;
    try
      StrWhere1 := MakeWhere;       // le MakeWhere renvoie les critères de sélection
      //TJA /FQ10819
      //ajout d'une table temporaire sur les suspectscompl
      ExecuteSQL('SELECT * INTO #SUSPECTSCOMPL FROM SUSPECTSCOMPL WHERE RSC_SUSPECT IN (SELECT RSU_SUSPECT FROM SUSPECTS LEFT JOIN SUSPECTSCOMPL ON RSC_SUSPECT=RSU_SUSPECT WHERE '+StrWhere1+')');


      if StrWhere1 <> '' then       //toute la base
      //TJA 07/05/2008 //FQ 10819
      //jointure avec SUSPECTSCOMPL
        ExecuteSQL('DELETE FROM SUSPECTSCOMPL WHERE RSC_SUSPECT IN (SELECT RSU_SUSPECT FROM SUSPECTS LEFT JOIN SUSPECTSCOMPL ON RSC_SUSPECT=RSU_SUSPECT WHERE '+StrWhere1+')')
      else
        ExecuteSQL('DELETE FROM SUSPECTSCOMPL WHERE RSC_SUSPECT IN (SELECT RSU_SUSPECT FROM SUSPECTS)');
      MoveCurProgressForm('Veuillez patienter');

      if not (GetCheckBoxState('ELEMCIBL') = cbUnchecked) then
      begin
        if StrWhere1 <> '' then       // toute la base
        //TJA 07/05/2008
        //FQ 10819
          ExecuteSQL('DELETE FROM CIBLAGEELEMENT WHERE RVB_SUSPECT IN (SELECT RSU_SUSPECT FROM SUSPECTS LEFT JOIN #SUSPECTSCOMPL ON RSC_SUSPECT=RSU_SUSPECT WHERE '+StrWhere1+')')
        else
          ExecuteSQL('DELETE FROM CIBLAGEELEMENT WHERE RVB_SUSPECT IN (SELECT RSU_SUSPECT FROM SUSPECTS)');
        MoveCurProgressForm('Veuillez patienter');
      end;

      if StrWhere1 <> '' then       // toute la base
        //TJA 07/05/2008
        //FQ 10819
        ExecuteSQL('DELETE FROM SUSPECTS WHERE RSU_SUSPECT IN (SELECT RSU_SUSPECT FROM SUSPECTS LEFT JOIN #SUSPECTSCOMPL ON RSC_SUSPECT=RSU_SUSPECT WHERE '+StrWhere1+')')
      else
        ExecuteSQL('DELETE FROM SUSPECTS');
      MoveCurProgressForm('Veuillez patienter');
      //suppression de la table temporaire
      ExecuteSQL('DROP TABLE #SUSPECTSCOMPL');
      COMMITTRANS;
    except
      on E: Exception do
        begin
          ROLLBACK;
          ExecuteSQL('DROP TABLE #SUSPECTSCOMPL');
        end;

    end;
  end;


  FiniMoveProgressForm;

  //F.Q.UpdateCriteres;
  F.BChercheClick(Ecran);

  F.FListe.ClearSelected;


end;


procedure TOF_RTSUPPRSUSP.FlisteOnDblClick(Sender: Tobject);
begin
  if VarIsNull(GetField('RSU_SUSPECT')) then
    exit;

  AGLLanceFiche('RT', 'RTSUSPECTS', '', GetField('RSU_SUSPECT'), 'MONOFICHE;ACTION=CONSULTATION');


end;


Initialization
  registerclasses ( [ TOF_RTSUPPRSUSP ] ) ;
end.
