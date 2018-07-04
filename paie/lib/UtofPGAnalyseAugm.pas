{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_AUGMENTATION ()
Mots clefs ... : TOF;saisie des augmentations de salaires
*****************************************************************}
Unit UtofPGAnalyseAugm ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,HDB,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry,PgOutils2,EntPaie,P5Def,HTB97,uTob,HSysMenu,Ed_Tools ,P5Util,LookUp,ParamSoc;

Type
  TOF_PGANALYSEAUGM = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure ExitEdit(Sender: TObject);
    Function RendRequeteResp(ValeurResp,ChampResp : String) : String;
  end ;

Implementation

procedure TOF_PGANALYSEAUGM.OnLoad ;
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

procedure TOF_PGANALYSEAUGM.OnArgument (S : String ) ;
var Num : Integer;
    Defaut : THEdit;
    Combo : THValComboBox;
    MultiC : THMultiValComboBox;
begin
  Inherited ;
        SetControlText('ANNEE',GetParamSoc('SO_PGAUGANNEE'));
        Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
        Defaut:=ThEdit(getcontrol('RESPONSABS'));
        SetControlCaption('LIBSALARIE','');
        SetControlCaption('LIBABSENCE','');
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
        For Num := 1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
         end;
         VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
         Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
end ;

procedure TOF_PGANALYSEAUGM.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

Function TOF_PGANALYSEAUGM.RendRequeteResp(ValeurResp,ChampResp : String) : String;
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
  registerclasses ( [ TOF_PGANALYSEAUGM ] ) ;
end.

