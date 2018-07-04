unit dpTOMOrga;
// TOM de la table DPORGA

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,HDB,UTOM,UTOB,HTB97,
      Spin,Menus,Windows, Messages,
      
{$IFDEF EAGLCLIENT}
      MaineAGL,
      eFiche,
{$ELSE}
      FE_Main, {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
      Fiche,
{$ENDIF}
      Graphics, AnnOutils, EventDecla;

type
  TOM_DPORGA = Class (TOM)
    procedure OnArgument(stArgument : String ) ; override ;
    procedure OnNewRecord  ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord  ; override ;
    procedure OnAfterUpdateRecord; override;
    procedure OnClose; override ;

    function  CalculNbTitre(GuidPerDos : String): Integer;
    //LM20070516 function  RecupNomGerant(GuidPerDos : String): String;
    procedure Onclickjur(Sender : TObject);
    //LM20070404 procedure OnExitDatefin(Sender : TObject);
    //LM20070404 procedure OnExitDuree(Sender : TObject);
  private
    EvDecla : TEventDecla ;//LM20070404
    sGuidPer_c       : string;
    TobAnnPer        : TOB;
    TobJuri          : TOB;
    DateCreation     : TDateTime;
    FormJur          : String;
    //RefGerant        : String;
    bCharge          : Boolean;
    Onglet           : String;
//    TSavDatefin      : TDateTime;
    //LM20070404 procedure CalculDuree;
    //LM20070404 procedure CalculDateExercice;
    //LM20070511 procedure ElipsisInterlocClick(Sender: TObject);
    function  RendLibelleContact(GuidPer: string): String;
    //LM20070412 procedure ElipsisPropriClick(Sender: TObject);
    procedure FormKeyDownOrga(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DOR_CONVENTIONNE_OnClick(Sender: TObject);
    procedure DOR_SECTIONBNC_OnClick(Sender: TObject);
    //LM20070511 procedure DOR_ETABLTS_OnClick(Sender: TObject);
    //LM20070511 procedure BT_ETABLISSEMENTS_OnClick(Sender: TObject);
    //LM20070404 procedure BLIENENTITE_OnClick(Sender: TObject);
    //LM20070404 procedure DOR_ATTACHEMENT_OnClick(Sender: TObject);
    //LM20070404 procedure BLIENPROPRIETAIRE1_OnClick(Sender: TObject);
    //LM20070404 procedure DOR_LOCAGERANCE_OnClick(Sender: TObject);
    //LM20070511 procedure DOR_PROBLEME_OnClick(Sender: TObject);
    procedure GereGrpboxMed;
    procedure GereConventionne;
    procedure GereBNC;
    //LM20070511 procedure GereEtablts;
    //LM20070404 procedure GereFiliales;
    //LM20070404 procedure GereLocagerance;
    //LM20070511 procedure GereParticularite;
    procedure setEventRessource (fld : string) ; //LMO20060901
    procedure ficheRessource (sender : TObject); //LMO20060901
    procedure BT_DIVERSWINNER_OnClick (Sender: TObject); //PGR 01/2007 Ajout bouton Divers dans l'onglet Winner
    procedure ShowJuridique (Sender: TObject);
  end ;


///////////// IMPLEMENTATION ////////////
implementation

uses dpOutils, galOutil, galSystem,
     UDossierSelect,
     DpJurOutils,
     paramsoc,
     EntDP; //LMO

procedure TOM_DPORGA.OnClose;
begin
  inherited;
  // EvDecla.Free ; MD20070705 assuré par le owner du TComponent
  if TobAnnPer<>Nil then TobAnnPer.Free ;
  if TobJuri<>Nil then TobJuri.Free ;
end;


procedure TOM_DPORGA.OnArgument(stArgument : String );
var POPMENUF11 : TPopUpMenu;
    i : Integer;
    Action : String;
begin
  inherited; // traitement des actions
  EvDecla := TEventDecla.create(Ecran) ;//+LM20070404
  EvDecla.Rebranche('TJUR_DATEDEBUTEX','OnClick', ShowJuridique);
  EvDecla.Rebranche('TJUR_DATEFINEX','OnClick', ShowJuridique); //-LM20070404

  bCharge := True;
  Action := ReadTokenSt(stArgument);
  Onglet := ReadTokenSt(stArgument);
  sGuidPer_c := ReadTokenSt(stArgument);
  if sGuidPer_c = '' then sGuidPer_c := TFFiche(Ecran).FLequel;

  Ecran.Caption := 'Organisation : ' + GetNomPer(sGuidPer_c);
  UpdateCaption(Ecran);

  // affectation des évènements
  POPMENUF11 := TPopUpMenu(getcontrol('POPMENUF11'));
  if POPMENUF11<>Nil then
    begin
    for i:=0 to POPMENUF11.Items.Count-1 do
      begin

      (*LM20070511 if POPMENUF11.Items[i].Name = 'ETA' then
          POPMENUF11.Items[i].OnClick := BT_ETABLISSEMENTS_OnClick*)
      (* //LM20070404 else if POPMENUF11.Items[i].Name = 'FIL' then
        POPMENUF11.Items[i].OnClick := BLIENENTITE_OnClick *)
      (* //LM20070404 else if POPMENUF11.Items[i].Name = 'PFC' then
        POPMENUF11.Items[i].OnClick := BLIENPROPRIETAIRE1_OnClick;*)
      end;
    end;
  //LM20070511 THEdit(GetControl('PERSONNEACONTACTER')).OnElipsisClick:=ElipsisInterlocClick;
  //LM20070412 THEdit(GetControl('NOMPROPRIETRAIRE')).OnElipsisClick:=ElipsisPropriClick;
  Ecran.OnKeyDown := FormKeyDownOrga;
  (*LM20070511 TCheckBox(GetControl('JUR_APPELPUB')).OnClick := Onclickjur;
  TCheckBox(GetControl('JUR_REDRESS')).OnClick := Onclickjur;
  TCheckBox(GetControl('JUR_LIQUID')).OnClick := Onclickjur; *)
  TCheckBox(GetControl('JUR_CAC')).OnClick := Onclickjur;
  //LM20070404 THEdit(GetControl('DOR_DATEFINEX')).Onexit := OnExitDatefin;
  //LM20070404 THEdit(GetControl('DOR_DUREE')).Onexit := OnExitDuree;
  //LM20070404 THEdit(GetControl('DOR_DATEDEBUTEX')).Onexit := OnExitDuree;
  THDBValComboBox(GetControl('DOR_SECTIONBNC')).OnClick := DOR_SECTIONBNC_OnClick;
  TCheckBox(GetControl('DOR_CONVENTIONNE')).OnClick := DOR_CONVENTIONNE_OnClick;
  //LM20070412 TCheckBox(GetControl('DOR_ETABLTS')).OnClick := DOR_ETABLTS_OnClick;
  //LM20070412 TToolBarButton97(GetControl('BT_ETABLISSEMENTS')).OnClick := BT_ETABLISSEMENTS_OnClick;
  //LM20070404 TToolBarButton97(GetControl('BLIENENTITE')).OnClick := BLIENENTITE_OnClick;
  //LM20070404 TCheckBox(GetControl('DOR_ATTACHEMENT')).OnClick := DOR_ATTACHEMENT_OnClick;
  //LM20070404 TToolBarButton97(GetControl('BLIENPROPRIETAIRE1')).OnClick := BLIENPROPRIETAIRE1_OnClick;
  //LM20070404 TCheckBox(GetControl('DOR_LOCAGERANCE')).OnClick := DOR_LOCAGERANCE_OnClick;
  //LM20070511 TCheckBox(GetControl('DOR_PROBLEME')).OnClick := DOR_PROBLEME_OnClick;
  //PGR 01/2007 Ajout bouton Divers dans l'onglet Winner
  if GetControl('BT_DIVERSWINNER')<>Nil then
    TToolBarButton97(GetControl('BT_DIVERSWINNER')).OnClick := BT_DIVERSWINNER_OnClick;

  // 20/06/01
  SetControlVisible('BINSERT', False);
  // #### 5/12/01 PROVISOIRE
  // On n'affiche pas l'établissement si on n'est pas sur le DP du cabinet,
  // en attendant fiche établissement qui sait taper dans la base dossier
  // car ici on écrasait toujours les établissements de la compta du cabinet !!!
  if V_PGI.DbName<>VH_Doss.DBSocName then
    begin
    SetControlVisible('BT_ETABLISSEMENTS', False);
    //SetControlEnabled('DOR_ETABLTS', False);
    SetControlEnabled('DOR_NBETABLTS', False);
    SetControlEnabled('TDOR_NBETABLTS', False);
    // TCheckBox(GetControl('DOR_ETABLTS')).OnClick := Nil;
    // TToolBarButton97(GetControl('BT_ETABLISSEMENTS')).OnClick := Nil;
    end;
  SetControlVisible('GRPMONNAIE', False);

  // #### Obligations : partie "Responsables par domaine" inadaptée (non liée GI)
  if not vh_dp.Group then SetControlVisible('GRPBOXRESP', False);//LMO
  if vh_dp.Group then //LMO20060901
  begin
    setEventRessource ('DOR_UTILRESPFISCAL_') ;
    setEventRessource ('DOR_UTILRESPCOMPTA_');
    setEventRessource ('DOR_UTILRESPJURID_');
    setEventRessource ('DOR_UTILRESPSOCIAL_');
  end ;

  SetControlVisible('GRPBOXCABINET', False);
  TGroupBox(GetControl('GRPBOXTRAITE')).Top := 10 ;

  // $$$ JP 27/09/05 - onglet Winner ne doit pas apparaitre si la gestion Winner n'est pas activée
  if GetParamSocSecur ('SO_MDLIENWINNER', FALSE) = FALSE then
  begin
       SetControlProperty ('PWINNER', 'TabVisible', FALSE);
  end
  else
  begin
       SetControlVisible ('DOR_WETABENTITE', FALSE);
       SetControlVisible ('TDOR_WETABENTITE', FALSE);
  end;

  if vh_dp.group then SetControlProperty ('PCABINET', 'TabVisible', False);  //LM20070511
end;

procedure TOM_DPORGA.ficheRessource (sender : TObject); //LMO20060901
var q : TQuery ;
    st,id, mode, usr : string ;
begin
  if sender=nil then exit ;
  id:=TControl(sender).name;
  delete(id,length(id),1) ;
  usr := getControlText(id);
  if usr<>'' then q:=openSql('select ARS_RESSOURCE from RESSOURCE where ARS_UTILASSOCIE = "' + usr + '"', true) ;
  if (usr<>'') and (not q.eof) then
  begin
    mode:='MODIFICATION';
    id := q.Fields[0].asString;
  end
  else
  begin
    mode:='CREATION';
    id:='';
  end ;
  if usr<>'' then ferme(q) ;
  st:=aglLanceFiche ('AFF', 'RESSOURCE', '', id, 'ACTION='+Mode + ';' ) ;
end ;

procedure TOM_DPORGA.setEventRessource (fld : string) ; //LMO20060901
var ctrl : TControl;
begin
    ctrl := getControl(fld) ;
    if ctrl <> nil then TButton(ctrl).OnClick := FicheRessource
                   else SetControlVisible(fld, False);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPORGA.OnNewRecord;
begin
  inherited;
  bCharge:=True;
  SetField('DOR_GUIDPER', sGuidPer_c);
  bCharge := False;
end;

procedure TOM_DPORGA.OnLoadRecord;
var
    Q            : TQuery;
    nbetab : Integer;
    //visib        : Boolean;
    //bDateEnabled : Boolean;
    st, select : string ;
begin
  bCharge := True;

  Ecran.Caption := 'Organisation : '+GetNomPer(sGuidPer_c);
  UpdateCaption(Ecran);

  // #### Corrections sur l'enreg en cours : voir comment c'est fait
  // dans dpTomFiscal/dpTomSocial, dans les procédures Initialisation
  // => elles évitent le msg "voulez-vous enregistrer les modifications ?"
  // car elles sont appelées dans OnArgument, sur l'enreg FLequel

//  TSavDatefin := iDate1900;
  (*LM20070404
  if GetField('DOR_DUREE') = 0 then
    begin ModeEdition(DS); SetField('DOR_DUREE', 12); end;
  *)

  (*  //LM20070404
  TobAnnPer.Free ;
  TobAnnPer:=TOB.Create('DPPERSO',Nil,-1) ;
  TobAnnPer.SelectDB('"'+sGuidPer_c+'"',Nil,TRUE) ;
  TobAnnPer.PutEcran(Ecran) ;
  *)

  ModeEdition(DS);//LM20070412 LM>MD on ne devrait pas faire ça mais le putEcran est fait sur le Field...

  TobJuri.Free ;
  TobJuri:=TOB.Create('JURIDIQUE',Nil,-1) ;
  //+LM20070404
  //TobJuri.SelectDB('"'+sGuidPer_c+'"',Nil,TRUE) ;
  select := FaitLeSelect(Ecran) ;
  st := 'select ANN_FORME, ANN_FORMEGEN, ' + select + ' '+
        'from ANNUAIRE A ' + #13 +
        'left outer join DPFISCAL F on F.DFI_GUIDPER = A.ANN_GUIDPER ' + #13 +
        'left outer join JURIDIQUE J on J.JUR_GUIDPERDOS = A.ANN_GUIDPER ' + #13 +
        'left outer join DPORGA O on O.DOR_GUIDPER = A.ANN_GUIDPER ' + #13 +
        'where ANN_GUIDPER = "'  + sGuidPer_c +'"' ;
  q:=opensql(st,true) ;
  if not q.Eof then
  begin
    FormJur := Q.FindField('ANN_FORMEGEN').AsString;
    if (FormJur = 'SC') then FormJur := q.FindField('ANN_FORME').AsString;
  end ;
  TobJuri.SelectDB('', q, true) ;
  ferme(q) ;
  TobJuri.PutEcran(Ecran) ;
  //-LM20070404

  // personne à contacter
  //mcd 20/02/2006 SetControlText('PERSONNEACONTACTER', RendLibelleContact(nodp));
  // SetControlText('PERSONNEACONTACTER', RendLibelleContact(GetField('DOR_GUIDPER')));

  // date début activité dans table juridique
  //+LM20070404 DateCreation :=GetDateCreation(sGuidPer_c);    centraliser dans un autre select unique + composant JUR_DEBACT dans la fiche
  st :=getControlText('JUR_DEBACT')  ;
  if isValidDate(st) then DateCreation := strToDate(st) else DateCreation:=idate1900; 
  //-LM20070404

  if FormJur <> 'ASS' then
    TTabSheet(GetControl('TSASSOCIATION')).TabVisible := FALSE;

  // groupe "Patrimoine de la sci"
  if FormJur <> 'SCI' then SetControlVisible ('gbSCI', FALSE);

  // Si le dossier a une base de prod !
  if DBExists ('DB' + VH_DOSS.NoDossier) then
    begin
    nbetab := NbEnreg(VH_Doss.DBSocName+'.dbo.ETABLISS', 'ET_ETABLISSEMENT', '');
    if GetField('DOR_NBETABLTS')<>nbetab then
      begin
      ModeEdition(DS);
      SetField('DOR_NBETABLTS', nbetab);
      if nbetab>=2 then SetField('DOR_ETABLTS', 'X')
      else SetField('DOR_ETABLTS', '-');
      end;
    end;

  // zones visibles/invisibles
  GereGrpboxMed;
  GereConventionne;
  GereBNC;
  //LM20070511 GereEtablts;
  //LM20070404 GereFiliales;
  //LM20070404 GereLocagerance;
  //LM20070511 GereParticularite;

  // $$$ JP 20/10/03: si dates existantes pour N dans DPTABCOMPTA, interdit de modifier les dates compta de la fiche DPORGA
  (*+LM20070404
  bDateEnabled := not ExisteSQL ('SELECT DTC_NODOSSIER FROM DPTABCOMPTA JOIN DOSSIER ON DTC_NODOSSIER=DOS_NODOSSIER AND DOS_GUIDPER="' + sGuidPer_c + '" WHERE DTC_MILLESIME="N"');
  SetControlEnabled ('DOR_DATEDEBUTEX', bDateEnabled);
  SetControlEnabled ('DOR_DATEFINEX', bDateEnabled);
  SetControlEnabled ('DOR_DUREE', bDateEnabled);

  // MD 05/06/06 - Mais on les initialise si c'est la première fois qu'on rentre dans l'enregistrement
  if DS.State in [dsInsert] then
    begin
    Q := OpenSQL('SELECT DTC_DATEDEB, DTC_DATEFIN, DTC_DUREE FROM DPTABCOMPTA, DOSSIER '
     + 'WHERE DOS_GUIDPER="' + sGuidPer_c + '" AND DOS_NODOSSIER=DTC_NODOSSIER AND DTC_MILLESIME="N"', True);
    If Not Q.Eof then
      begin
      SetField('DOR_DATEDEBUTEX', Q.FindField('DTC_DATEDEB').AsDateTime);
      SetField('DOR_DATEFINEX',   Q.FindField('DTC_DATEFIN').AsDateTime);
      SetField('DOR_DUREE',       Q.FindField('DTC_DUREE').AsInteger);
      end;
    Ferme(Q);

    Q := OpenSQL('SELECT DTC_DUREE FROM DPTABCOMPTA, DOSSIER '
     + 'WHERE DOS_GUIDPER="' + sGuidPer_c + '" AND DOS_NODOSSIER=DTC_NODOSSIER AND DTC_MILLESIME="N-"', True);
    If Not Q.Eof then
      begin
      SetField('DOR_DUREEPREC',   Q.FindField('DTC_DUREE').AsInteger);
      end;
    Ferme(Q);
    end;

  // $$$ JP 22/10/2003: idem pour durée exercice N-1 si enreg. existant dans DPTABCOMPTA pour "N-"
  bDateEnabled := not ExisteSQL ('SELECT DTC_NODOSSIER FROM DPTABCOMPTA JOIN DOSSIER ON DTC_NODOSSIER=DOS_NODOSSIER AND DOS_GUIDPER="' + sGuidPer_c + '" WHERE DTC_MILLESIME="N-"');
  SetControlEnabled ('DOR_DUREEPREC', bDateEnabled);
  -LM20070404 *)

  // onglet demandé
  if (Onglet<>'') and (TTabSheet(GetControl(Onglet)).TabVisible) then
    TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl(Onglet));

  bCharge := False;
end;


procedure TOM_DPORGA.OnUpdateRecord;
var
    //Tfin : TDateTime;
    //Tdeb,Tprec : TDateTime;
    //Annee, Mois, Jour: Word;
    //ChMois : String;
    bCac_l : boolean;
begin
  (* //LM20070404
  Tfin := StrToDate(GetField('DOR_DATEFINEX'));
  Tdeb := StrToDate(GetField('DOR_DATEDEBUTEX'));
  if Not (Tdeb < Tfin) then
    begin
    ErreurChamp('DOR_DATEFINEX', 'PGeneral', Self, 'La date de fin d''exercice doit être supérieure à la date de début') ;
    exit;
    end  ;
  if GetField ('DOR_DUREE') < 0 then
    begin
    ErreurChamp('DOR_DUREE', 'PGeneral', Self, 'La durée ne peut pas être négative') ;
    exit;
    end;
  if GetField ('DOR_DUREE') > 24 then
    begin
    ErreurChamp('DOR_DUREE', 'PGeneral', Self, 'La durée est au maximum de 24 mois') ;
    exit;
    end;
  if (Tdeb < DateCreation) AND (DateCreation <> iDate1900) then
    begin
    ErreurChamp('DOR_DATEDEBUTEX', 'PGeneral', Self, 'La date de début d''exercice doit être antérieure à la date de création de l''entreprise') ;
    exit;
    end;
  Tprec := PlusMois ( Tdeb , -1*(GetField ('DOR_DUREEPREC')) );
  if (Tprec < DateCreation) then
    begin
    ErreurChamp('DOR_DATEDEBUTEX', 'PGeneral', Self, 'La date de début d''exercice précédent doit être antérieure à la date de création de l''entreprise') ;
    exit;
    end;
*)

  //pour association
  if (FormJur = 'ASS') then
    begin
    if GetField ('DOR_FORMEASSO') = '' then
      begin
      ErreurChamp('DOR_FORMEASSO', 'TSASSOCIATION', Self, 'La forme de l''association est obligatoire') ;
      exit;
      end;
    if GetField ('DOR_CATEGASSO') = '' then
      begin
      ErreurChamp('DOR_CATEGASSO', 'TSASSOCIATION', Self, 'La catégorie de l''association est obligatoire') ;
      exit;
      end;
    if GetField ('DOR_SECTEURASSO') = '' then
      begin
      ErreurChamp('DOR_SECTEURASSO', 'TSASSOCIATION', Self, 'Le secteur de l''association est obligatoire') ;
      exit;
      end;
    if GetField ('DOR_TYPEASSO') = '' then
      begin
      ErreurChamp('DOR_TYPEASSO', 'TSASSOCIATION', Self, 'Le type de l''association est obligatoire') ;
      exit;
      end;

    if Not ((GetCheckBoxState('DOR_ADHECOTI')=cbChecked)
         or (GetCheckBoxState('DOR_DONMANUEL')=cbChecked)
         or (GetCheckBoxState('DOR_DONNATURE')=cbChecked)
         or (GetCheckBoxState('DOR_VTEPUBLI')=cbChecked)
         or (GetCheckBoxState('DOR_BIENFAIS')=cbChecked)
         or (GetCheckBoxState('DOR_DONATION')=cbChecked)
         or (GetCheckBoxState('DOR_SUBVENTION')=cbChecked)
         or (GetCheckBoxState('DOR_BILLETERIE')=cbChecked)
         or (GetCheckBoxState('DOR_REMBTPARMBS')=cbChecked)
         ) then
      begin
      ErreurChamp('DOR_ADHECOTI', 'TSASSOCIATION', Self, 'Il est obligatoire de renseigner au moins une nature de ressources pour l''association') ;
      exit;
      end;
    end; // Fin association

  if (GetCheckBoxState('DOR_CONVENTIONNE')=cbChecked)
   and (FormJur = 'IND') then
    begin
    if GetField('DOR_VALEURSNIR') = 0 then
      begin
      ErreurChamp('DOR_VALEURSNIR', 'TSBNC', Self, 'La valeur du SNIR est obligatoire pour un praticien conventionné') ;
      exit;
      end;
    end;

  // purges éventuelles
  if (FormJur <> 'IND') then
    SetField('DOR_SECTIONBNC', '');

  (* DOR_REGLEFISC
  if (GetField('DOR_REGLEBNC')='') and (GetControlText('DFI_REGLEFISC')='BNC') then
    // REC = recettes dépenses, CRD = créances dettes
    SetField('DOR_REGLEBNC', 'RED');
  *)

  // maj chp DOR_DUREE
  //LM20070404 CalculDuree;

  // en cas de nouvel enreg ?
  if tobjuri <>Nil then //mcd 13/06/07 pour Ok veriftob
    TobJuri.PutValue('JUR_GUIDPERDOS',GetField('DOR_GUIDPER'));

  (*LM20070511 if TCheckBox(GetControl('JUR_APPELPUB')).checked then
    TobJuri.PutValue('JUR_APPELPUB',  'X')
  else
    TobJuri.PutValue('JUR_APPELPUB',  '-');
  if TCheckBox(GetControl('JUR_REDRESS')).Checked then
    TobJuri.PutValue('JUR_REDRESS',  'X')
  else
    TobJuri.PutValue('JUR_REDRESS',  '-');
  if TCheckBox(GetControl('JUR_LIQUID')).Checked then
    TobJuri.PutValue('JUR_LIQUID',  'X')
  else
    TobJuri.PutValue('JUR_LIQUID',  '-');*)

  if tobjuri <>Nil then //mcd 13/06/07 pour Ok veriftob
  begin
   bCac_l := (TobJuri.GetString('JUR_CAC') = 'X');

  if bCac_l <> TCheckBox(GetControl('JUR_CAC')).Checked then
  begin
     if TCheckBox(GetControl('JUR_CAC')).Checked then
     begin
       TobJuri.PutValue('JUR_CAC',  'X');
     end
     else
     begin
       TobJuri.PutValue('JUR_CAC',  '-');
       SupprimeAttach(GetField('DOR_GUIDPER'), 'CCT', 'CCS');
     end;
  end;

  if (TobJuri.ExistDB) then
     TobJuri.UpdateDB(FALSE)
  else
    begin
    // MD 27/07/01 - ne pas écraser si dossier juridique existe
    TobJuri.PutValue('JUR_CODEDOS', '&#@');
    TobJuri.PutValue('JUR_NOORDRE', 1);
    TobJuri.PutValue('JUR_TYPEDOS', 'STE');
    // fin MD
    TobJuri.InsertOrUpdateDB(FALSE);
    end;
  end;// fin mcd 13/06/07
  //--- Mise a jour du mois de cloture de la fiche annuaire
  (* //LM20070404
  DecodeDate(GetField('DOR_DATEFINEX'), Annee, Mois, Jour);
  ChMois:=Format ('%2.2d',[Mois]);
  ExecuteSQL ('UPDATE ANNUAIRE SET ANN_MOISCLOTURE="'+ChMois+'" WHERE ANN_GUIDPER="'+ GetField ('DOR_GUIDPER') +'"');
  *)
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPORGA.OnAfterUpdateRecord;
begin
   inherited;
   if IsInside(Ecran) then
      ReloadTomInsideAfterInsert(TFFiche(Ecran), DS, ['DOR_GUIDPER'], [sGuidPer_c]);

   //--- FQ 11781 Organisation : Rêgles de comptabilisation
   if (GetField('DOR_REGLEBNC')='CRD') or (GetField('DOR_REGLEBNC')='RED') or (GetField('DOR_REGLEBNC')='TRE') then
     if (DBExists ('DB'+VH_DOSS.NoDossier)) then
     begin
       if (GetField('DOR_REGLEBNC')='CRD') then
         ExecuteSQL ('UPDATE DB'+VH_DOSS.NoDossier+'.DBO.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_CPTVARECDEP"')
       else
         ExecuteSQL ('UPDATE DB'+VH_DOSS.NoDossier+'.DBO.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_CPTVARECDEP"');
     end;
end;

(* //LM20070404
procedure TOM_DPORGA.CalculDuree;
// calcul de la durée de l'exercice par différence fin-deb
var
  Tfin : TDateTime;
  Tdeb : TDateTime;
  iNbmois : Integer;
begin
  Tdeb := StrToDate(GetField('DOR_DATEDEBUTEX'));
  Tfin := StrToDate(GetField('DOR_DATEFINEX'));

  // Le calcul pourrait être fait directement avec Tfin-Tdeb...
  iNbmois := 0;
  While (Tdeb <= Tfin) do
    begin
    iNbmois := iNbmois+1;
    Tdeb := PlusMois ( Tdeb , 1 );
    end;
  if GetField ('DOR_DUREE') <> iNbmois then
    SetField('DOR_DUREE', iNbmois);
end;
*)

(*//LM20070404
procedure TOM_DPORGA.CalculDateExercice;
var Tdeb :TDateTime;
    Tfin :TDateTime;
    wA, wM, wJ : word;
begin
  if (GetField ('DOR_DUREE') <> 0) then
    begin
    Tdeb := StrToDate(GetField ('DOR_DATEDEBUTEX'));
    Tdeb := Tdeb;
    Tfin := PlusMois ( Tdeb , GetField('DOR_DUREE'));
    Tfin := Tfin-1;
    if StrToDate(GetField ('DOR_DATEFINEX')) <> Tfin then
      begin
      SetField ('DOR_DATEFINEX', Tfin);
//      TSavDatefin := Tfin;
      DecodeDate(StrToDate(GetField ('DOR_DATEFINEX')), wA, wM, wJ);
      SetField ('DOR_EXERCICE', inttostr(wA));
      end;
    end;
end;
*)

procedure TOM_DPORGA.OnChangeField(F: TField);
begin
{
if ((F.FieldName = 'DOR_DATEFINEX') AND (DS.State in [dsInsert,dsEdit])) then
  //LM20070404 CalculDuree;

if ((F.FieldName = 'DOR_DUREE') AND (DS.State in [dsInsert,dsEdit])) then
  CalculDateExercice;
}

  (* //LM20070404
  if (F.FieldName = 'DOR_DATEFINEX') and (DS.State in [dsInsert,dsEdit]) then
    begin
    if ((StrToDate(GetField ('DOR_DATEDEBUTEX')) <> iDate1900))
     and ((StrToDate(GetField ('DOR_DATEFINEX')))<> iDate1900) then
      begin
      if not (StrToDate(GetField ('DOR_DATEDEBUTEX')) < StrToDate(GetField ('DOR_DATEFINEX'))) then
        begin
        TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('PGeneral'));
        SetFocusControl('DOR_DATEFINEX');
        PGIInfo('La date de fin d''exercice doit être supérieure à la date de début', TitreHalley) ;
        end  ;
      end;
    end;
  *)
  (*//LM20070404
  if (F.FieldName = 'DOR_ATTACHEMENT')
   and Not (GetCheckBoxState('DOR_ATTACHEMENT')=cbChecked)
   and (DS.State in [dsInsert,dsEdit]) then
    SetField ('DOR_NBRATTACH',  0);
  *)
end;


function TOM_DPORGA.CalculNbTitre(GuidPerDos : String): Integer;
begin
  // inutile car n'y passait jamais : Result := GetField ('DOR_NBRATTACH');
  Result := NbEnreg('ANNULIEN', '*', 'WHERE ANL_GUIDPERDOS="'+GuidPerDos+'"'
         +' AND ANL_FONCTION="FIL"');
end;

{LM20070516
function TOM_DPORGA.RecupNomGerant(GuidPerDos : String): String;
// affecte aussi le no RefGerant
var
    Q1 : TQuery;
begin
  Result := '';
  RefGerant := '';
  Q1 :=  OpenSQL('SELECT ANL_NOMPER, ANL_GUIDPER FROM ANNULIEN WHERE ANL_GUIDPERDOS="'
          +GuidPerDos+'" AND ANL_FONCTION="PFC"', True);
  if (Q1=nil) or (Q1.EOF) then
    SetControlText ('NOMPROPRIETRAIRE', '')
  else
    begin
    Result := Q1.Fields[0].AsString;
    RefGerant := Q1.Fields[1].AsString;
    end;
  Ferme(Q1);
end;
}

{LM20070511
procedure TOM_DPORGA.ElipsisInterlocClick(Sender: TObject);
begin
  ModeEdition(DS);
  // $$$ JP 11/10/05 - nature YY et non plus DP
  //mcd 12/2005 AGLLanceFiche ('YY', 'ANNUINTERLOC_PER',IntToStr(GetField('DOR_NODP')),'','ACTION=MODIFICATION');
  AccesContact ( GetField ('DOR_GUIDPER') );
  SetControlText ('PERSONNEACONTACTER', RendLibelleContact (GetField ('DOR_GUIDPER')));
end;}


(* //LM20070412
procedure TOM_DPORGA.ElipsisPropriClick(Sender: TObject);
var ValChamp : String;
    Nb       : Integer;
begin
  ModeEdition(DS);
  Nb := NbEnreg('ANNULIEN', '*', 'WHERE ANL_GUIDPERDOS="'+GetField('DOR_GUIDPER')+'"'
         +' AND ANL_FONCTION="PFC"');
  ValChamp := RecupNomGerant(GetField('DOR_GUIDPER'));
  if (ValChamp = '') or (Nb > 1) then
    begin
    // plusieurs propriétaires => affiche la liste pour en choisir un à afficher
    ValChamp := AGLLanceFiche('DP','LIENINTERPROP',
      'ANL_GUIDPERDOS='+GetField('DOR_GUIDPER')+';ANL_FONCTION=PFC','',
      GetField('DOR_GUIDPER'));
    if (ValChamp = '') then ValChamp := RecupNomGerant(GetField('DOR_GUIDPER'));
    end
  else
    // 1 seul prop, affiche directement sa fiche
    AGLLanceFiche('YY','ANNUAIRE',RefGerant,RefGerant,'ACTION=MODIFICATION')  ;

  if (ValChamp <> '') then SetControlText('NOMPROPRIETRAIRE', ValChamp);
end;
*)

function TOM_DPORGA.RendLibelleContact(GuidPer: String): String;
var Q : TQuery;
begin
  Result := '';
 (* mcd 12/2005  Q := OpenSQL('select ANI_NOM, ANI_PRENOM from ANNUINTERLOC where ANI_CODEPER = '
      +InttoStr(nodp)+' and ANI_PRINCIPAL = "X"', TRUE);  *)
  Q := OpenSQL('select C_NOM, C_PRENOM from CONTACT where C_GUIDPER ="'
      +GuidPer+'" and C_PRINCIPAL = "X"', TRUE);
  if not Q.eof then
    Result := Q.Fields[0].AsString + ' ' + Q.Fields[1].AsString;
  Ferme(Q);
end;


procedure TOM_DPORGA.FormKeyDownOrga(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  POPMENUF11 : TPopUpMenu;
begin
  POPMENUF11 := TPopUpMenu(getcontrol('POPMENUF11'));
  if (POPMENUF11 <> nil) and (TPageControl(GetControl('Pages')).ActivePage = TTabSheet(GetControl('PGeneral'))) then
    begin
    if ((ORD(Key) = VK_F11) or (ORD(Key) = VK_RIGHT)) then
      if POPMENUF11.items.Count >  0 then
//        POPMENUF11.Popup (Mouse.CursorPos.x-100,Mouse.CursorPos.x-100);
        POPMENUF11.Popup (+100,+100{Mouse.CursorPos.x});
    end;
  inherited;
end;


procedure TOM_DPORGA.Onclickjur (Sender : TObject);
// pour passer en mode edit quand on modifie une zone qui ne dépend pas
// de la table principale (DPORGA) mais d'une tob sur une autre table (JURIDIQUE)
begin
  // évite le passage lors du TobJuri.PutEcran du OnLoadRecord
  if bCharge then exit;
  if not(DS.State in [dsInsert,dsEdit]) then
    begin
    ModeEdition(DS);
    // renseigne un chp bidon sinon pas de msg "Voulez-vous enregistrer"
    SetField('DOR_REGLEBNC', GetControlText('DOR_REGLEBNC'));
    end;
end;

(*//LM20070404
procedure TOM_DPORGA.OnExitDatefin(Sender : TObject);
begin
  CalculDuree;
end;
*)

(* //LM20070404
procedure TOM_DPORGA.OnExitDuree(Sender : TObject);
begin
  CalculDateExercice;
end;
*)


procedure TOM_DPORGA.DOR_SECTIONBNC_OnClick(Sender: TObject);
begin
  GereGrpboxMed;
end;


procedure TOM_DPORGA.DOR_CONVENTIONNE_OnClick(Sender: TObject);
begin
  GereConventionne;
end;


(*LM20070511
procedure TOM_DPORGA.DOR_ETABLTS_OnClick(Sender: TObject);
begin
  GereEtablts;
end;*)

(*LM20070511
procedure TOM_DPORGA.BT_ETABLISSEMENTS_OnClick(Sender: TObject);
var nbetab: Integer;
begin
  AGLLanceFiche('YY', 'YYETABLISS', '', '', '');
  nbetab := NbEnreg(VH_Doss.DBSocName+'.dbo.ETABLISS', 'ET_ETABLISSEMENT', '');
  THDBSpinEdit(GetControl('DOR_NBETABLTS')).value := nbetab;
  SetControlChecked('DOR_ETABLTS', nbetab >= 2);
  GereEtablts;
end;
*)

(*//LM20070404
procedure TOM_DPORGA.BLIENENTITE_OnClick(Sender: TObject);
var nbfil: Integer;
begin
  AGLLanceFiche('DP','LIENINTERFILIALE','ANL_GUIDPERDOS='+GetField('DOR_GUIDPER')+';'
               +'ANL_FONCTION=FIL','',GetField('DOR_GUIDPER')+';FIL');
  nbfil := CalculNbTitre(GetField('DOR_GUIDPER'));
  SetControlText('DOR_NBRATTACH', IntToStr(nbfil));
  SetField('DOR_NBRATTACH', nbfil);
end;
*)

(*//LM20070404
procedure TOM_DPORGA.DOR_ATTACHEMENT_OnClick(Sender: TObject);
var res: Integer;
begin
  // évite le passage lors du chargement
  if bCharge then exit;

  // on décoche "Filiales" alors qu'il y a des filiales
  if (GetCheckBoxState('DOR_ATTACHEMENT')<>cbChecked)
   and (THDBSpinEdit(GetControl('DOR_NBRATTACH')).Value <> 0) then
    begin
    res := SupprimeAttach(GetField('DOR_GUIDPER'), 'FIL', '');
    if res = 0 then
      SetField('DOR_NBATTACH', res)
    else
      SetControlChecked('DOR_ATTACHEMENT', TRUE);
    end;
  GereFiliales;
end;
*)
(* //LM20070404
procedure TOM_DPORGA.BLIENPROPRIETAIRE1_OnClick(Sender: TObject);
begin
  // on remplit aussi param Lequel pour remplir les critères du mul
  AGLLanceFiche('DP','LIENINTERPROP','ANL_GUIDPERDOS='+GetField('DOR_GUIDPER')
    +';ANL_FONCTION=PFC','',GetField('DOR_GUIDPER')+';PFC');
  SetControlText('NOMPROPRIETRAIRE', RecupNomGerant(GetField('DOR_GUIDPER')));
end;
*)
(*//LM20070404
procedure TOM_DPORGA.DOR_LOCAGERANCE_OnClick(Sender: TObject);
var res: Integer;
begin
  // évite le passage lors du chargement
  if bCharge then exit;

  // si on décoche alors qu'il y a un propriétaire
  if (GetCheckBoxState('DOR_LOCAGERANCE')<>cbChecked) then
    begin
    res := 0;
    if RecupNomGerant(GetField('DOR_GUIDPER'))<>'' then
      res := SupprimeAttach(GetField('DOR_GUIDPER'), 'PFC', '');
    // si on n'a pas pu le supprimer, on recoche
    if res <> 0 then SetControlChecked('DOR_LOCAGERANCE', True);
    end;
  // affichages
  GereLocagerance;
end;
*)

(* LM20070511
procedure TOM_DPORGA.DOR_PROBLEME_OnClick(Sender: TObject);
begin
  GereParticularite;
end;*)


procedure TOM_DPORGA.GereGrpboxMed;
// gère le groupe Activité médicale
var sectionbnc: String;
    visib : Boolean;
begin
  sectionbnc := THDBValComboBox(GetControl('DOR_SECTIONBNC')).Value;
  visib := (sectionbnc = 'MED') or (sectionbnc = 'DEN')
        or (sectionbnc = 'AUX') or (sectionbnc = 'SAG');
  SetControlVisible('GRPBOXMED', visib);
end;


procedure TOM_DPORGA.GereConventionne;
// gère les zones praticien conventionné dans l'onglet BNC
var actif: Boolean;
begin
  actif := TCheckBox(GetControl('DOR_CONVENTIONNE')).checked;
  SetControlEnabled('DOR_VALEURSNIR', actif);
  SetControlEnabled('DOR_DATESNIR', actif);
  SetControlEnabled('DOR_ORIGINESNIR', actif);
  SetControlEnabled('DOR_VALSNIRPREC', actif);
  SetControlEnabled('TDOR_VALEURSNIR', actif);
  SetControlEnabled('TDOR_DATESNIR', actif);
  SetControlEnabled('TDOR_ORIGINESNIR', actif);
  SetControlEnabled('TDOR_VALSNIRPREC', actif);
end;

procedure TOM_DPORGA.GereBNC;
begin
  // $$$ JP 26/04/06 - FQ 10914: on tient compte QUE de DFI_REGLEFISC
  TTabSheet(GetControl('TSBNC')).TabVisible := (GetControlText ('DFI_REGLEFISC') = 'BNC'); //LM20070404 GetField ('DOR_REGLEFISC') = 'BNC';

  // si on est en exploitant individuel au régime BNC
{  if (FormJur = 'IND') and (GetControlText('DFI_REGLEFISC') = 'BNC') then
      TTabSheet(GetControl('TSBNC')).TabVisible := TRUE
  else
      TTabSheet(GetControl('TSBNC')).TabVisible := FALSE;}
end;


(*LM20070511 
procedure TOM_DPORGA.GereEtablts;
var actif : Boolean;
begin
  actif := (GetCheckBoxState('DOR_ETABLTS')=cbChecked)
   and (THDBSpinEdit(GetControl('DOR_NBETABLTS')).Value >= 2);
  // seuls les établissements du cabinet sont consultables
  // car fiche YYETABLISS liée à table ETABLISS locale
  SetControlEnabled('BT_ETABLISSEMENTS',
      (actif) and (V_PGI.DBName=VH_Doss.DBSocName)
    );
  SetControlEnabled('DOR_NBETABLTS', actif);
  SetControlEnabled('TDOR_NBETABLTS', actif);
end;
*)
(*//LM20070404
procedure TOM_DPORGA.GereFiliales;
var actif : Boolean;
begin
  actif := TCheckBox(GetControl('DOR_ATTACHEMENT')).checked ;
  SetControlEnabled('BLIENENTITE', actif);
  SetControlEnabled('DOR_NBRATTACH', actif);
  SetControlEnabled('TDOR_NBRATTACH', actif);
  SetControlVisible('DOR_COMPTESCONSO', actif);
end;
*)

(* //LM20070404
procedure TOM_DPORGA.GereLocagerance;
var actif: Boolean;
begin
  actif := (GetCheckBoxState('DOR_LOCAGERANCE')=cbChecked);
  SetControlEnabled('BLIENPROPRIETAIRE1', actif);
  SetControlEnabled('TNOMPROPRETAIRE', actif);
  SetControlEnabled('NOMPROPRIETRAIRE', actif);
  SetControlText('NOMPROPRIETRAIRE', RecupNomGerant(GetField('DOR_GUIDPER')));
end;
*)

(*LM20070511
procedure TOM_DPORGA.GereParticularite;
var actif : Boolean;
begin
  actif := (GetCheckBoxState('DOR_PROBLEME')=cbChecked);
  // rq : la zone DOR_TYPEPROBLEME n'est pas utilisée,
  // remplacée par ces zones venant du juridique :
  SetControlEnabled('JUR_APPELPUB', actif);
  SetControlEnabled('JUR_REDRESS', actif);
  SetControlEnabled('JUR_LIQUID', actif);
end;*)

//PGR 01/2007 Ajout bouton Divers dans l'onglet Winner
procedure TOM_DPORGA.BT_DIVERSWINNER_OnClick(Sender: TObject);
begin
  if sGuidPer_c <> '' then
    AglLanceFiche('DP', 'DIVERS_WINNER', sGuidPer_c ,sGuidPer_c ,'ACTION=MODIFICATION');
end;

procedure TOM_DPORGA.ShowJuridique (Sender: TObject);
//Var ModalResult: TModalResult ;
begin
  (* Correction agl en attente...
  AGLLanceFiche('JUR','JURIDIQUE', sGuidPer_c, sGuidPer_c, 'ACTION=MODIFICATION;' + sGuidPer_c +
                ';;;;;TabSheet=PVIESOCIALE;','', ModalResult) ;
  *)
//  AGLLanceFiche('JUR','JURIDIQUE', sGuidPer_c, sGuidPer_c, 'ACTION=MODIFICATION;' + sGuidPer_c +
//                ';;;;;TabSheet=PVIESOCIALE;') ;

   //*** BM (29/01/2008) :
   //*** Accès dossier juridique depuis bureau -> LanceFicheDP
   //*** Création auto si n'existe pas

  LanceFicheDP('JUR', 'JURIDIQUE', sGuidPer_c + ';STE;1', sGuidPer_c + ';STE;1',
               'ACTION=MODIFICATION;' + sGuidPer_c + ';STE;1;DP;;TabSheet=PVIESOCIALE;', 'JURIDIQUE', 'JUR');
  //if ModalResult=mrOk then OnLoadRecord ;
end ;

Initialization
  registerclasses([TOM_DPORGA]) ;
end.
