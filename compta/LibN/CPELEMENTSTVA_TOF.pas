{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/09/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPELEMENTSTVA ()
Mots clefs ... : TOF;CPELEMENTSTVA
*****************************************************************}
Unit CPELEMENTSTVA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Vierge,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB,
     HTB97,
     AGLInit,
     Ent1,
     SaisUtil,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     fe_main,
{$ENDIF}
     ParamSoc;

Type
  TOF_CPELEMENTSTVA = Class (TOF)
    EMP_JOURNAL : THValComboBox ;
    BFerme : TButton ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure BFermeClick (Sender: TObject);
   private
    FRetour : String;

  end ;

function CPLanceFiche_ParamTVA ( vStParam: string = '') : String ;

Implementation

function CPLanceFiche_ParamTVA ( vStParam: string = '') : String ;
begin
   Result := AGLLanceFiche('CP','CPELEMENTSTVA','','', vstParam);
end;

procedure TOF_CPELEMENTSTVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPELEMENTSTVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPELEMENTSTVA.OnUpdate ;
var
  T : TOB;
  CodeTaux, Numero : String;
  Q : TQuery;
begin

  T := TOB.Create ('', nil, -1);
  CodeTaux:=THValComboBox(GetControl('TAUXTVA')).Value;
  {Recherche du CC_LIBRE pour savoir quel ECHEENC on alimente}
  Numero:=RechDom('TTTVA', CodeTaux, False, 'AND CC_LIBRE>0 AND CC_LIBRE<5',  True, '');
  {FQ22230  23.01.08  YMO On prend le taux 1 pour ECHEDEBIT}
  If CodeTaux='' then
  begin
    Q :=  OpenSQL('Select CC_CODE from CHOIXCOD Where CC_TYPE="TX1" AND CC_LIBRE="1"', FALSE);
    if not Q.EOF then
    CodeTaux := Q.FindField('CC_CODE').AsString;
    ferme(Q);
    Numero:='0';
  end;

  T.AddChampSupValeur('CODETVA', CodeTaux);
  T.AddChampSupValeur('NUMTAUX', Numero);

  {REGFISCAL  EXIGIBILITE  TAUXTVA}

  TheTOB := T;

  Ecran.Close ;
end ;

procedure TOF_CPELEMENTSTVA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPELEMENTSTVA.OnArgument (S : String ) ;
var
  RegFiscal, Exigibilite : String;
  Q : Tquery;
begin

 // THValComboBox(GetControl('TAUXTVA')).Plus:='AND CC_LIBRE>0 AND CC_LIBRE<5';

{REGFISCAL  EXIGIBILITE  TAUXTVA}
  Regfiscal := ReadTokenSt(S);
  Exigibilite := ReadTokenSt(S);

  If RegFiscal<>'X' then
    THValComboBox(GetControl('REGFISCAL')).Itemindex := THValComboBox(GetControl('REGFISCAL')).Values.IndexOf(RegFiscal)
  else
    THValComboBox(GetControl('REGFISCAL')).Itemindex := 0;

  If Exigibilite<>'X' then
    THValComboBox(GetControl('EXIGIBILITE')).Itemindex := THValComboBox(GetControl('EXIGIBILITE')).Values.IndexOf(Exigibilite)
  else
    THValComboBox(GetControl('EXIGIBILITE')).Itemindex := 0;

  {Taux 1 proposé par défaut}
  Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="TX1" AND CC_LIBRE=1',True) ;
  if Not Q.EOF then
    THValComboBox(GetControl('TAUXTVA')).ItemIndex:=THValComboBox(GetControl('TAUXTVA')).Values.IndexOf(Q.FindField('CC_CODE').AsString);
  Ferme(Q) ;

  TButton(GetControl('BFerme')).OnClick:=BFermeClick;

  TheTOB := nil;

  Inherited ;
end ;

procedure TOF_CPELEMENTSTVA.OnClose ;
begin
  TFVierge(Ecran).Retour:=FRetour;
  Inherited ;
end ;

procedure TOF_CPELEMENTSTVA.BFermeClick (Sender: TObject);
begin
  FRetour:='S';
end;

Initialization
  registerclasses ( [ TOF_CPELEMENTSTVA ] ) ;
end.
