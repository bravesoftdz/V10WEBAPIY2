{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_AUGMENTATION ()
Mots clefs ... : TOF;saisie des augmentations de salaires
*****************************************************************}
Unit UTofPGMulAugmIntegration ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,HDB,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry,PgOutils2,EntPaie,P5Def,HTB97,uTob,HSysMenu,Ed_Tools ,P5Util,ParamSoc;

Type
  TOF_PGMULAUGMINTEGRE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Argument : String;
    procedure GrilleDblClick (Sender : TObject);
    procedure ExitEdit(Sender: TObject);
    Function RendRequeteResp(ValeurResp,ChampResp : String) : String;
  end ;
  var
        Q_MulAug : THQuery;
       {$IFNDEF EAGLCLIENT}
        ListeAug : THDBGrid ;
        {$ELSE}
        ListeAug : THGrid ;
        {$IFDEF EAGLCLIENT}
       TFMAug : TFMul;
      {$ENDIF}
        {$ENDIF}
        TobAug : Tob;

Implementation

procedure TOF_PGMULAUGMINTEGRE.OnLoad ;
var StWhere,Annee,WhereResp : String;
    CRespVar : THMultiValComboBox;
begin
  Inherited ;
        StWhere := '';
        If GetControlText('ANNEE') <> '' then
        begin
             Annee := RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False);
             If StWhere = '' then StWhere := '(PBG_ANNEE="'+Annee+'")'
             else StWhere := StWhere + 'AND (PBG_ANNEE="'+Annee+'")';
        end;
        CRespVar := THMultiValComboBox(GetControl('RESPONSVAR'));
        If (GetControlText('RESPONSVAR') <> '') and (CRespVar.tous = False) then
        begin
             WhereResp := RendRequeteResp(GetControlText('RESPONSVAR'),'PSE_RESPONSVAR');
             If StWhere = '' then StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+')'
             else  StWhere := StWhere + 'AND PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+')';
        end;
        If (GetControlText('RESPONSABS') <> '') then
        begin
             If StWhere = '' then StWhere := 'PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSABS="'+GetControlText('RESPONSABS')+'")'
             else  StWhere := StWhere + 'AND PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSABS="'+GetControlText('RESPONSABS')+'")';
        end;
        SetControlText('XX_WHERE',StWHere);
end ;

procedure TOF_PGMULAUGMINTEGRE.OnArgument (S : String ) ;
var Num : Integer;
    Defaut : THEdit;
begin
  Inherited ;
        SetControlCaption('LIBSALARIE','');
        SetControlCaption('LIBABSENCE','');
        Argument := ReadTokenPipe(S,';');
        {$IFNDEF EAGLCLIENT}
        ListeAug := THDBGrid(GetControl('FListe'));
        {$ELSE}
        ListeAug := THGrid(GetControl('FListe'));
        TFMAug := TFMul(Ecran);
        {$ENDIF}
        SetControlText('ANNEE',GetParamSoc('SO_PGAUGANNEE'));
        SetControlText('CSAISIE','');
        Q_MulAug := THQuery(Ecran.FindComponent('Q'));
        Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
        If defaut <> Nil then Defaut.OnExit := ExitEdit;
        Defaut:=ThEdit(getcontrol('RESPONSABS'));
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PPBG__TRAVAILN'+IntToStr(Num)),GetControl ('TPBG__TRAVAILN'+IntToStr(Num)));
         end;
         If ListeAug <> Nil Then ListeAug.OnDblClick := GrilleDblClick ;
         VisibiliteStat (GetControl ('PBG_CODESTAT'),GetControl ('TPBG_CODESTAT')) ;
end ;

procedure TOF_PGMULAUGMINTEGRE.GrilleDblClick (Sender : TObject);
var Where : String;
    Pages : TPageControl;
begin
       Pages := TPageControl(GetControl('Pages'));
        if (ListeAug = nil) then Exit;
        if (ListeAug.NbSelected = 0) and (not ListeAug.AllSelected) then
        begin
             MessageAlerte('Aucun élément sélectionné');
             exit;
        end;
        Where := RecupWhereCritere(Pages);
        AGLLanceFiche('PAY','AUGM_INTEGRA','','',Where);
        SetControlProperty('BSelectAll','Down',False);
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULAUGMINTEGRE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

Function TOF_PGMULAUGMINTEGRE.RendRequeteResp(ValeurResp,ChampResp : String) : String;
var StResp,Requete : String;
begin
     Result := '';
     requete := '';
     While ValeurResp <> '' do
     begin
          StResp := ReadTokenPipe(ValeurResp,';');
          If Requete <> '' then requete := Requete + ' OR ' + ChampResp+'="'+StResp+'"'
          else Requete := ChampResp+'="'+StResp+'"';
     end;
     Result := Requete;
end;

Initialization
  registerclasses ( [ TOF_PGMULAUGMINTEGRE ] ) ;
end.

