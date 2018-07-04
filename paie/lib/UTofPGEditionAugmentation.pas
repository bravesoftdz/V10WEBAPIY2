{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITAUGM ()
Mots clefs ... : TOF;PGEDITAUGM
*****************************************************************}
Unit UTofPGEditionAugmentation ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     QRS1,
{$else}
     eMul,
     uTob,
     eQRS1,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     P5Util,
     P5Def,
     Paramsoc,
     PGOutils2,
     EntPaie,
     HQRY,
     UTOF ;

Type
  TOF_PGEDITAUGM = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure ExitEdit(Sender: TObject);
    procedure ChangeAnnee(Sender : TObject);
  end ;

Implementation

procedure TOF_PGEDITAUGM.OnUpdate ;
var Pages : TPageControl;
    Where,SQL : String;
begin
  Inherited ;
        Pages := TPageControl(GetControl('PAGES'));    
        Where := RecupWhereCritere(Pages);
        SQL := 'SELECT * FROM BUDGETPAIE LEFT JOIN SALARIES ON PSA_SALARIE=PBG_SALARIE '+Where;
        TFQRS1(Ecran).WhereSQL := SQL;
end ;

procedure TOF_PGEDITAUGM.OnArgument (S : String ) ;
var Num : Integer;
    Defaut : THEdit;
    Combo : THValComboBox;
begin
  Inherited ;
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  Defaut:=ThEdit(getcontrol('PBG_SALARIE'));
  Defaut.OnExit := ExitEdit;
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num >4 then Break;
    VisibiliteChampSalarie (IntToStr(Num),GetControl ('PPBG__TRAVAILN'+IntToStr(Num)),GetControl ('TPBG__TRAVAILN'+IntToStr(Num)));
  end;
  VisibiliteStat (GetControl ('PBG_CODESTAT'),GetControl ('TPBG_CODESTAT'));
  Combo := THValComboBox(GetControl('ANNEE'));
  If Combo <> Nil then Combo.OnChange := ChangeAnnee;
  SetControlText('ANNEE',GetParamSoc('SO_PGAUGANNEE'));
end ;

procedure TOF_PGEDITAUGM.ExitEdit(Sender: TObject);
var edit : thedit;
begin
    edit:=THEdit(Sender);
    if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITAUGM.ChangeAnnee(Sender : TObject);
begin
  SetControlText('PBG_ANNEE',RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False));
end;

Initialization
  registerclasses ( [ TOF_PGEDITAUGM ] ) ;
end.

