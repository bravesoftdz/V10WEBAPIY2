{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 16/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_AUGMENTATION ()
Mots clefs ... : TOF;saisie des augmentations de salaires
*****************************************************************}
Unit UTofPGMul_AugmGenere ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,FE_Main,HDB,
{$else}
     eMul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry,PgOutils2,EntPaie,P5Def,HTB97,uTob,HSysMenu,Ed_Tools ,P5Util,ParamSoc;

Type
  TOF_PGMULAUGMGENERE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Argument : String;
    procedure GrilleDblClick (Sender : TObject);
    procedure ExitEdit(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;
  var
        Q_MulAug : THQuery;
        {$IFNDEF EAGLCLIENT}
        ListeAug : THDBGrid ;
        {$ELSE}
        ListeAug : THGrid ;
        TFMAug : TFMul;
        {$ENDIF}
        TobAug : Tob;

Implementation

procedure TOF_PGMULAUGMGENERE.OnLoad ;
var StWhere,Annee,StDateArret : String;
    dateArret : TDateTime;
begin
  Inherited ;
        StWhere := '';
        Annee := RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False);
        If GetControlText('CGENERE') <> '' then
        begin
                If GetControlText('CGENERE') = 'NONGENERE' then StWhere := 'PSA_SALARIE NOT IN (SELECT PBG_SALARIE FROM BUDGETPAIE WHERE PBG_ANNEE="'+Annee+'" )'
                else if GetControlText('CGENERE') = 'GENERE'  then StWhere := 'PSA_SALARIE IN (SELECT PBG_SALARIE FROM BUDGETPAIE WHERE PBG_ANNEE="'+Annee+'" )';
        end;
        if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
        Begin
             DateArret:=StrtoDate(GetControlText('DATEARRET'));
             StDateArret:=' (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
             StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
        End
        else StDateArret:='';
        If StWhere <> '' then StWhere := StWhere + ' AND '+StDateArret
        else StWhere := StDateArret;
        SetControlText('XX_WHERE',StWHere);
end ;

procedure TOF_PGMULAUGMGENERE.OnArgument (S : String ) ;
var Num : Integer;
    Defaut : THEdit;
    Check : TCheckBox;
begin
  Inherited ;
        Argument := ReadTokenPipe(S,';');
        {$IFNDEF EAGLCLIENT}
        ListeAug := THDBGrid(GetControl('FListe'));
        {$ELSE}
        ListeAug := THGrid(GetControl('FListe'));
        TFMAug := TFMul(Ecran);
        {$ENDIF}
        Q_MulAug := THQuery(Ecran.FindComponent('Q'));
        Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
        If Defaut<>nil then Defaut.OnExit:=ExitEdit;
        For Num := 1 to VH_Paie.PGNbreStatOrg do    
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
         end;
         If ListeAug <> Nil Then ListeAug.OnDblClick := GrilleDblClick ;
         VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
         SetControlText('ANNEE', GetParamSoc('SO_PGAUGANNEE'));
         SetControlvisible('DATEARRET',True);
         SetControlvisible('TDATEARRET',True);
         SetControlEnabled('DATEARRET',False);
         SetControlEnabled('TDATEARRET',False);
         SetControlChecked('CKSORTIE',True);
         SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
         SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
         Check:=TCheckBox(GetControl('CKSORTIE'));
         if Check <> Nil then Check.OnClick:=OnClickSalarieSortie;
end ;

procedure TOF_PGMULAUGMGENERE.GrilleDblClick (Sender : TObject);
var Retour,Annee,Where : String;
    Pages : TPageControl;
    Bt : TToolBarButton97;
begin
       If GetControlText('ANNEE') = '' then
       begin
            PGIBox('Vous devez sélectionner une année',Ecran.Caption);
            Exit;
       end;
        Pages := TPageControl(GetControl('PAGES'));
        if (ListeAug = nil) then Exit;
        If Q_MulAug.RecordCount < 1 then
        begin
                PGIBox('Aucun salarié ne correspond aux critères',ecran.Caption);
                Exit;
        end;
        if (ListeAug.NbSelected = 0) and (not ListeAug.AllSelected) then
        begin
                MessageAlerte('Vous devez sélectionner le(s) salarié(s) à saisiir');
                exit;
        end;
        Annee := RechDom('PGANNEESOCIALE',GetControlText('ANNEE'),False);
        Where := RecupWhereCritere(Pages);
        Retour := AGLLanceFiche('PAY','AUGMGENERE','','',Annee+';'+Where);
        If Retour = 'Maj' then
        begin
             ListeAug.ClearSelected;
             ListeAug.AllSelected:=False;
             Bt  :=  TToolbarButton97 (GetControl('BCherche'));
             if Bt  <>  NIL then Bt.click;
        end;
end;

procedure TOF_PGMULAUGMGENERE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGMULAUGMGENERE.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

Initialization
  registerclasses ( [ TOF_PGMULAUGMGENERE ] ) ;
end.

