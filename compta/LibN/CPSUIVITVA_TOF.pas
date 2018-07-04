{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSUIVITVA ()
Mots clefs ... : TOF;CPSUIVITVA
*****************************************************************}
Unit CPSUIVITVA_TOF ;

Interface

Uses
     Htb97,
     TofMeth,
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,         // AGLLanceFiche
{$ELSE}
     eMul,
     MainEAgl,        // AGLLanceFiche
{$ENDIF}
{$IFDEF COMPTA}
     Saisie,		    // Pour Saisie eAGL
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Ent1,		              // Pour EstMonaieIn et GetPeriode
     Ed_Tools,              // Pour le videListe
     TiersPayeur,           // Pour les fonctions xxxTP
     HStatus,               // Pour la barre d'état
     AGLInit,               // TheMulQ
     Zcompte,
     SaisComm,
     dialogs,              // Pour les procédures de MAJ des comptes
     HSysMenu,
     ULibExercice,
     SaisBase,
     CPELEMENTSTVA_TOF,
     saisutil,
     uTob,
     paramsoc;

Type
  TOF_CPSUIVITVA = class(TOF_Meth)
   // Eléments interface
    E_NATUREPIECE    : THValComboBox ;
    E_EXERCICE       : THValComboBox ;
    E_ETABLISSEMENT  : THValComboBox ;
    E_GENERAL        : THEdit ;
    E_DATECOMPTABLE  : THEdit ;
    E_DATECOMPTABLE_ : THEdit ;
    E_DATECREATION   : THEdit ;
    E_DATECREATION_  : THEdit ;
    E_DATEECHEANCE   : THEdit ;
    E_DATEECHEANCE_  : THEdit ;
    COLLECTIF        : THEdit ;
    COLLECTIF_       : THEdit ;
    AUXILIAIRE       : THEdit ;
    AUXILIAIRE_      : THEdit ;
    NUMEROPIECE      : THEdit ;
    NUMEROPIECE_     : THEdit ;
    REFINTERNE       : ThEdit ;
    CPTEXCEPT        : THEdit ;
    XX_WHERE         : THEdit ;
    EDATESITUATION   : THEdit ;
    EXERCICE         : THValComboBox ;
    BSUIVANT         : TToolbarButton97;
    BVALIDER         : TToolbarButton97;
    BVALIDER1        : TToolbarButton97;
    CBRETEFFETESC    : TCheckBox;
    CBENSITUATION    : TCheckBox;
    REGTVA           : THMultiValComboBox ;
    NATUREGENE       : THMultiValComboBox ;
    CODETVA          : THMultiValComboBox ;
    MsgBox    : THMsgBox;
  {$IFDEF EAGLCLIENT}
    FListe : THGrid ;
  {$ELSE}
    FListe : THDBGrid ;
  {$ENDIF}
    FEcran    : TFMul ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure ValideClick (Sender: TObject);
    procedure NatureGeneOnChange (Sender : TObject) ;
    procedure ExerciceOnChange (Sender : TObject) ;
    procedure BSuivantClick(Sender: TObject);
    procedure CBENSITUATIONClick(Sender: TObject);
    procedure EDATESITUATIONChange (Sender : TObject) ;
    procedure ParamRuptures ;
    procedure ParamTitresCols ;
    function  GetCriteres : string ;
    
  end ;

function CPLanceFiche_SUIVITVA( vStParam : string = '' ) : string ;

Implementation

uses
//  RepDevEur, // ChangeLeTauxDevise
  UlibEcriture,
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , Constantes
  ,uLibWindows; // TestJoker, TraductionTHMultiValComboBox

function CPLanceFiche_SUIVITVA( vStParam : string = '' ) : string ;
begin
  Result := AGLLanceFiche('CP', 'CPSUIVITVA', '', '', vStParam);
end;

procedure TOF_CPSUIVITVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVITVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVITVA.OnUpdate ;
begin
  Inherited ;
end ;

function TOF_CPSUIVITVA.GetCriteres : string ;
var
    where, stTemp : string;
    etable,regime : String;
    listeCptExcept  : String ;
    stCpt, stMulti, stWhereMulti : String ;
begin
   where := '';

    // Recuperation des critere pour le XX_WHERE

   where := where + ' E_DATECOMPTABLE>="' + usDateTime(StrToDate(GetControlText('DATECOMPTABLE'))) + '" AND E_DATECOMPTABLE<="' + usdatetime(StrToDate(GetControlText('DATECOMPTABLE_'))) + '"';   //

   if (not (etable = '')) then where := where + ' AND E_ETABLISSEMENT="' + etable + '"';
   if (not (regime = '')) then where := where + ' AND E_REGIMETVA="' + regime + '"';

   where := where + ' ' + stTemp ;

   // Condition sur les collectifs :
   where := where + ' AND ' + ConvertitCaractereJokers(COLLECTIF, COLLECTIF_, 'E_GENERAL' );
   where := where + ' AND ' + ConvertitCaractereJokers(AUXILIAIRE, AUXILIAIRE_, 'E_AUXILIAIRE' );
   where := where + ' AND ' + ConvertitCaractereJokers(NUMEROPIECE, NUMEROPIECE_, 'E_NUMEROPIECE' );
   where := where + ' AND E_REFINTERNE LIKE"' +REFINTERNE.Text+'%"';

   // Comptes à exclure (Les comptes peuvent être séparés par des ',' ou des ';')
   if CPTEXCEPT.Text <> '' then
   begin
      listeCptExcept := FindEtReplace(CPTEXCEPT.Text,',',';', True);

      While (listeCptExcept <> '') do
      begin
        stCpt := Trim(ReadTokenSt(listeCptExcept)) ;

        if stCpt <> '' then
            where := where + ' AND E_AUXILIAIRE NOT LIKE "' + stCpt + '%"';
      end ;
   end;

    stMulti := REGTVA.Value;

    if stMulti <> '' then
    begin
      stWhereMulti := '';
      while stMulti <> '' do
        if stWhereMulti = ''	then
            stWhereMulti := ' AND E_REGIMETVA IN ("' + ReadTokenSt(stMulti) + '"'
        else
            stWhereMulti := stWhereMulti + ',"' + ReadTokenSt(stMulti) + '"';

      where := where + stWhereMulti + ')';
    end;

    stMulti := NATUREGENE.Value;

    if stMulti <> '' then
    begin
      stWhereMulti := '';
      while stMulti <> '' do
        if stWhereMulti = ''	then
            stWhereMulti := ' AND G_NATUREGENE IN ("' + ReadTokenSt(stMulti) + '"'
        else
            stWhereMulti := stWhereMulti + ',"' + ReadTokenSt(stMulti) + '"';

      where := where + stWhereMulti + ')';
    end;

    stMulti := CODETVA.Value;

    if stMulti <> '' then
    begin
      stWhereMulti := '';
      while stMulti <> '' do
        if stWhereMulti = ''	then
            stWhereMulti := ' AND E_TVA IN ("' + ReadTokenSt(stMulti) + '"'
        else
            stWhereMulti := stWhereMulti + ',"' + ReadTokenSt(stMulti) + '"';

      where := where + stWhereMulti + ')';
    end;


  // ====================
  // ==== LE LETTRAGE ===
  // ====================

  // Condition sur le lettrage
  if TCheckBox(GetControl('TRAITESLETTREES')).Checked=false then
  begin
      if CBENSITUATION.Checked then
        where := where + ' AND ( (E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI")'
                         + ' OR (E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX>"'+ USDateTime(StrToDate(EDATESITUATION.Text)) + '"))'
      else
        where := where + ' AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI"';
  end
  else  {retraitement des effets escomptés non échus}
  begin
     where := where + ' AND (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL")'
                    + ' AND J_EFFET="X" AND G_EFFET="X"'
                    + ' AND MP_CATEGORIE="LCR"';         {Modes de paiement type traite}
     if CBENSITUATION.Checked then
        where := where + ' AND E_DATEPAQUETMAX<="'+ USDateTime(StrToDate(EDATESITUATION.Text))
                       + '" AND E_DATEECHEANCE>"'+USDateTime(StrToDate(EDATESITUATION.Text))+'"';
  end;

  //where := where + ' E_DATEECHEANCE>="' + usDateTime(StrToDate(GetControlText('FDATE'))) + '" AND E_DATEECHEANCE<="' + usdatetime(StrToDate(GetControlText('FDATE_'))) + '"';

  {YMOO
  if not (Deductible) then
  begin
      where := where + ' AND (T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC" OR T_NATUREAUXI IS NULL)';
     // where := where + ' AND ((E_NATUREPIECE="RF" OR E_NATUREPIECE="OD" OR E_NATUREPIECE="OF") AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CAI")'

  end
  else
  begin
      where := where + ' AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD" OR T_NATUREAUXI IS NULL)';
    //  where := where + ' AND ((E_NATUREPIECE="RC" OR E_NATUREPIECE="OD" OR E_NATUREPIECE="OC") AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CAI")'

  end;}

  // lettrage non vide
  where := where + ' AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU<>"CLO" AND E_ECRANOUVEAU<>"OAN" ' ;

  // retour :
  result := where;

end ;


procedure TOF_CPSUIVITVA.OnLoad ;
begin
  Inherited ;

  ParamTitresCols ;

  SetControlText('XX_WHERE',GetCriteres);

end ;

procedure TOF_CPSUIVITVA.OnArgument (S : String ) ;
begin
  Inherited ;

  E_ETABLISSEMENT  := THValComboBox(GetControl('ETABLISSEMENT', True)) ;
  E_DATECOMPTABLE  := THEdit(GetControl('DATECOMPTABLE', True))  ;
  E_DATECOMPTABLE_ := THEdit(GetControl('DATECOMPTABLE_', True)) ;
  COLLECTIF        := THEdit(GetControl('COLLECTIF'));
  COLLECTIF_       := THEdit(GetControl('COLLECTIF_'));
  AUXILIAIRE       := THEdit(GetControl('AUXILIAIRE'));
  AUXILIAIRE_      := THEdit(GetControl('AUXILIAIRE_'));
  CPTEXCEPT        := THEdit(GetControl('CPTEXCEPT'));
  NUMEROPIECE      := THEdit(GetControl('NUMEROPIECE'));
  NUMEROPIECE_     := THEdit(GetControl('NUMEROPIECE_'));
  REFINTERNE       := THEdit(GetControl('REFINTERNE'));
  EDATESITUATION   := THEdit(GetControl('EDATESITUATION'));
  REGTVA           := THMultiValComboBox(GetControl('REGTVA'));
  NATUREGENE       := THMultiValComboBox(GetControl('NATUREGENE'));
  CODETVA          := THMultiValComboBox(GetControl('CODETVA'));
  BSUIVANT         := TToolBarButton97(GetControl('BSUIVANT'));
  CBRETEFFETESC    := TCheckBox(GetControl('CBRETEFFETESC'));
  CBENSITUATION    := TCheckBox(GetControl('CBENSITUATION'));
  EXERCICE         := THValComboBox(GetControl('EXERCICE_'));
  BVALIDER1        := TToolBarButton97(GetControl('BVALIDER1',true));

  BVALIDER1.OnClick   := ValideClick;
  BSUIVANT.OnCLick    := BSuivantClick;
  NATUREGENE.OnChange := NatureGeneOnChange;
  EXERCICE.OnChange   := ExerciceOnChange;

  CBENSITUATION.OnClick := CBENSITUATIONClick;
  EDATESITUATION.OnChange:=EDATESITUATIONChange;

  FEcran := TFMul(Ecran) ;

  If CBRETEFFETESC.Checked then SetControlVisible('BSUIVANT', True);

  If CBENSITUATION.Checked then
  begin
    EXERCICE.ItemIndex:=0;      {19.12.2007 YMO}
    E_DATECOMPTABLE.Text  := '01/01/1900';
    E_DATECOMPTABLE_.text := '01/01/2099';
    SetControlEnabled('EXERCICE_', False);
    SetControlEnabled('DATECOMPTABLE', False);
    SetControlEnabled('DATECOMPTABLE_', False);
  end;

  SetControlText('EDATESITUATION', DateToStr(VH^.EnCours.Fin));

  {racines de comptes}
  SetControlText('RACINECOC',GetParamSocSecur('SO_COLLCLIENC',''));
  SetControlText('RACINECOF',GetParamSocSecur('SO_COLLFOUENC',''));

  {longueur des racines de comptes}
  SetControlText('SZRACINECOC',IntToStr(Length(GetParamSocSecur('SO_COLLCLIENC',''))));
  SetControlText('SZRACINECOF',IntToStr(Length(GetParamSocSecur('SO_COLLFOUENC',''))));

end ;


procedure TOF_CPSUIVITVA.ValideClick(Sender: TObject);
var
  COLLCliParam, COLLFouParam : String;
begin
  If CBRETEFFETESC.Checked then
  begin
    BSuivant.Visible:=True;
    BValider1.Visible:=False;
  end;

  COLLCliParam:=GetParamSocSecur('SO_COLLCLIENC','');
  COLLFouParam:=GetParamSocSecur('SO_COLLFOUENC','');
  If ((Copy(COLLECTIF.Text, 1, Length(COLLCliParam))<>COLLCliParam)
     And (Copy(COLLECTIF.Text, 1, Length(COLLFouParam))<>COLLFouParam))
     Or ((Copy(COLLECTIF_.Text, 1, Length(COLLCliParam))<>COLLCliParam)
     And (Copy(COLLECTIF_.Text, 1, Length(COLLFouParam))<>COLLFouParam)) then
    MessageAlerte('Les racines de comptes ne correspondent pas au paramètrage TVA, voulez-vous continuer ?');
  BVALIDER := TToolBArButton97(GetControl('BVALIDER',true));

  BVALIDER.OnClick(BVALIDER);

end;


procedure TOF_CPSUIVITVA.NatureGeneOnChange (Sender : TObject) ;
var
  SQL : string;
  Lib : string;
begin

  SetControlText('CPTEXCEPT', '');

  {Constitution de la clause plus}
  TraductionTHMultiValComboBox(NatureGene, SQL, Lib, 'G_NATUREGENE');

  COLLECTIF.Plus := ' AND ' + SQL;
  COLLECTIF_.Plus := ' AND ' + SQL;

  {19.12.2007 YMO}
  TraductionTHMultiValComboBox(NatureGene, SQL, Lib, 'T_NATUREAUXI');

  SQL := UpperCase(StringReplace(SQL, 'COC', 'CLI',[rfReplaceAll])) ;
  SQL := UpperCase(StringReplace(SQL, 'COF', 'FOU',[rfReplaceAll])) ;
  SQL := UpperCase(StringReplace(SQL, 'COS', 'SAL',[rfReplaceAll])) ;
  SQL := UpperCase(StringReplace(SQL, 'COD', 'DIV',[rfReplaceAll])) ;

  AUXILIAIRE.Plus := ' AND ' + SQL;
  AUXILIAIRE_.Plus := ' AND ' + SQL;

end;

procedure TOF_CPSUIVITVA.ExerciceOnChange(Sender: TObject);
begin
   inherited;

   ExoToDates(EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_) ;

   if EXERCICE.ItemIndex = 0 then
   begin
        E_DATECOMPTABLE.Text  := '01/01/1900';
        E_DATECOMPTABLE_.text := '01/01/2099';
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 19/10/2007
Modifié le ... :   /  /
Description .. : retraitement des effets escomptés non échus
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSUIVITVA.BSuivantClick(Sender: TObject);
begin

  SetControlChecked('TRAITESLETTREES', True);

  BVALIDER := TToolBArButton97(GetControl('BVALIDER',true));

  BVALIDER.OnClick(BVALIDER);

  BSuivant.Visible:=False;
  BValider1.Visible:=True;

  SetControlChecked('TRAITESLETTREES', False);
end;

procedure TOF_CPSUIVITVA.CBENSITUATIONClick(Sender: TObject);
begin
  If CBENSITUATION.Checked then
  begin
    EXERCICE.ItemIndex:=0;      {19.12.2007 YMO}
    E_DATECOMPTABLE.Text  := '01/01/1900';
    E_DATECOMPTABLE_.text := '01/01/2099';
    SetControlEnabled('EXERCICE_', False);
    SetControlEnabled('DATECOMPTABLE', False);
    SetControlEnabled('DATECOMPTABLE_', False);
  end
  else
  begin
    SetControlEnabled('EXERCICE_', True);
    SetControlEnabled('DATECOMPTABLE', True);
    SetControlEnabled('DATECOMPTABLE_', True);
  end;
end;

procedure TOF_CPSUIVITVA.EDATESITUATIONChange (Sender : TObject) ;
begin
  SetControlText('DATECOMPTABLE_',EDATESITUATION.Text);
end;

procedure TOF_CPSUIVITVA.OnClose ;
begin
  Inherited ;
end ;


procedure TOF_CPSUIVITVA.ParamRuptures ;
begin
end;

procedure TOF_CPSUIVITVA.ParamTitresCols ;
var
  i : integer;
  Q : Tquery;
  SQL : String;
begin

if TCheckBox(GetControl('TRAITESLETTREES')).Checked=false then
   SetControlText( 'TITREETAT', 'Suivi de la TVA par les soldes' )
else
   SetControlText( 'TITREETAT', 'Effets escomptés non échus au '+EDATESITUATION.text ) ;

SetControlText( 'TITRECOL1', 'Acompte TTC' ) ;

for i:=1 to 4 do
begin
  SQL:='SELECT * FROM CHOIXCOD WHERE CC_TYPE="TX1" AND CC_LIBRE='+inttostr(i);
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then SetControlText( 'TITRECOL'+inttostr(i+1) , Q.FindField('CC_LIBELLE').AsString ) ;
  Ferme(Q) ;
end;


end;

Initialization
  registerclasses ( [ TOF_CPSUIVITVA ] ) ;
end.
