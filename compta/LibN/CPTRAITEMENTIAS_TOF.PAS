{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPTRAITEMENTIAS ()
Mots clefs ... : TOF;CPTRAITEMENTIAS
*****************************************************************}
Unit CPTRAITEMENTIAS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_main,
{$ENDIF}
     uTob,
     HEnt1,
     Ent1,               // ExoToDates
     HMsgBox,
     UTOF,
     ULibIAS14           // Type ParamIAS14 + Traitement
      ;

Procedure CPLanceFiche_TraitementIAS14( vStRange, vStLequel, vStArgs : String ) ;

Type
  TOF_CPTRAITEMENTIAS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    IAS14Process : TTraitementIAS14 ;
    procedure AvecExoCloClick   (Sender : TObject) ;
    procedure ExerciceChange    (Sender : TObject) ;
    procedure ChargeParamIAS14 ;
    Procedure InitParamIAS14 ;
    Procedure InitInfosIAS14 ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  ParamSoc           // GetParamSocSecur
  ;

Procedure CPLanceFiche_TraitementIAS14( vStRange, vStLequel, vStArgs : String ) ;
begin
  AGLLanceFiche('CP', 'CPTRAITEMENTIAS', vStRange, vStLequel, vStArgs) ;
end ;



procedure TOF_CPTRAITEMENTIAS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPTRAITEMENTIAS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPTRAITEMENTIAS.OnUpdate ;
var lInErr : Integer ;
begin
  Inherited ;

  // Chargement des paramètres
  ChargeParamIAS14 ;

  // Vérification des paramètres
  lInErr := IAS14Process.VerifieParamIAS14 ;
  if lInErr <> IAS_PASERREUR then
    begin
    IAS14Process.AfficheMessage ( lInErr,  Ecran.Caption ) ;
    // focus sur la zone qui pose pb :
    Case lInErr of
      IAS_ERRNOAXE        : SetFocusControl('AXE') ;
      IAS_ERRNOEXO        : SetFocusControl('EXERCICE') ;
      IAS_ERRDATEDEBUTFIN : SetFocusControl('DATECOMPTABLE') ;
      IAS_ERRNOPERDANSEXO : SetFocusControl('DATECOMPTABLE') ;
      end ;
    exit ;
    end ;

  // Confirmation
  if PgiAsk('Confirmez-vous le traitement des IAS 14 ?', Ecran.Caption) <> mrYes
    then Exit ;

  // Traitement
  lInErr := IAS14Process.TraitementIAS14 ;
  IAS14Process.AfficheMessage ( lInErr, Ecran.Caption ) ;

  // MAJ info dernier traitement
  InitInfosIAS14 ;

end ;

procedure TOF_CPTRAITEMENTIAS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPTRAITEMENTIAS.OnArgument (S : String ) ;
begin
  Inherited ;

  // SG6 26.01.05 Gestion mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    SetControlVisible('AXE',False);
    SetControlVisible('TAXE',False);
  end;


  // helpcontext
  Ecran.HelpContext := 7249500 ;

  // Affectation des évènements
  TCheckBox( GetControl('AVECEXOCLO', True) ).OnClick    := AvecExoCloClick ;
  THValComboBox( GetControl('EXERCICE', True) ).OnChange := ExerciceChange ;

  // Liste des natures de pièce limitée
  SetControlProperty( 'NATUREPIECE', 'PLUS', ' AND CO_CODE IN ("AF","FF","OD","FC","AC") ' ) ;

  // Init Param + infos
  InitParamIAS14 ;
  InitInfosIAS14 ;

  // Instanciation du process de traitement
  IAS14Process := TTraitementIAS14.Create ;

end ;

procedure TOF_CPTRAITEMENTIAS.OnClose ;
begin
  Inherited ;
  // Libération du process de traitement
  IAS14Process.Free;
end ;

procedure TOF_CPTRAITEMENTIAS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPTRAITEMENTIAS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPTRAITEMENTIAS.AvecExoCloClick(Sender: TObject);
begin
  if GetControlText('AVECEXOCLO') = 'X' then
    begin
    SetControlProperty('EXERCICE','DATATYPE','TTEXERCICE') ;
    SetControlText('EXERCICE','') ;
    IAS14Process.AfficheMessage( IAS_WNGEXOCLO, Ecran.Caption ) ;
    end
  else
    begin
    SetControlProperty('EXERCICE','DATATYPE','TTEXERCICEOUV') ;
    SetControlText('EXERCICE','') ;
    end ;

end;

procedure TOF_CPTRAITEMENTIAS.ExerciceChange(Sender: TObject);
begin
  ExoToDates( GetControlText('EXERCICE'), GetControl('DATECOMPTABLE'), GetControl('DATECOMPTABLE_') );
end;

procedure TOF_CPTRAITEMENTIAS.ChargeParamIAS14 ;
begin

  With IAS14Process do
    begin
    AvecExoClo         := GetControlText('AVECEXOCLO')='X' ;
    Exercice           := GetControlText('EXERCICE') ;
    DateDebut          := StrToDate( GetControlText('DATECOMPTABLE') ) ;
    DateFin            := StrToDate( GetControlText('DATECOMPTABLE_') ) ;
    if Not(VH^.AnaCroisaxe) then   Axe := GetControlText('AXE');

    if THMultiValComboBox( GetControl('NATUREGENE') ).Tous
      then NatureGene         := 'COF;COC;COS;COD;DIV;TIC;TID'
      else NatureGene         := Trim( GetControlText('NATUREGENE') ) ;
    if NatureGene[ length(NatureGene) ] = ';' then
      NatureGene := Copy(NatureGene, 1, length(NatureGene)-1 ) ;

    if THMultiValComboBox( GetControl('NATUREPIECE') ).Tous
      then NaturePiece        := 'AF;FF;OD;FC;AC'
      else NaturePiece        := Trim( GetControlText('NATUREPIECE') ) ;
    if NaturePiece[ length(NaturePiece) ]=';' then
      NaturePiece := Copy(NaturePiece, 1, length(NaturePiece)-1 ) ;

    Tolerance          := True ;
    ToleranceMontant   := Abs ( Valeur( GetControlText('TOLERANCEMONTANT') ) ) ;
    end ;

end;

procedure TOF_CPTRAITEMENTIAS.InitParamIAS14;
begin

  // Sans les eco clo par défaut
  SetcontrolText  ( 'AVECEXOCLO', '-' ) ;
  AvecExoCloClick ( nil ) ;

  // Exercice + Dates ( Exo d'entrée par défaut )
  SetControltext ( 'EXERCICE', VH^.Entree.Code ) ;
  ExerciceChange ( nil );

  // Axe ( 1er par défaut )
  SetControlText ( 'AXE', 'A1');

  // Nature de comptes ( Tous par défaut )
  SetControlText ( 'NATUREGENE', 'COF;COC;COS;COD;TIC;TID' ) ;

  // Nature de pièces ( Tous par défaut )
  SetControlText ( 'NATUREPIECE', 'AF;FF;OD;FC;AC' ) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 23/03/2005
Modifié le ... : 23/03/2005
Description .. : 
Suite ........ : 23/03/2005 : SBO Les zones dates sont laissées à blanc si 
Suite ........ : aucun traitement n'a encore eu lieu.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTRAITEMENTIAS.InitInfosIAS14;
var lDtDate : TDateTime ;
begin
  // Infos dernier traitement
  lDtDate := GetParamSocSecur('SO_CPIAS14DEBUT', iDate1900 ) ;
  if lDtDate <= iDate1900
    then SetControlText('DATELASTPER', '' )
    else SetControlText('DATELASTPER', DateToStr( lDtDate ) ) ;

  lDtDate := GetParamSocSecur('SO_CPIAS14FIN', iDate1900 ) ;
  if lDtDate <= iDate1900
    then SetControlText('DATELASTPER_', '' )
    else SetControlText('DATELASTPER_', DateToStr( lDtDate ) ) ;
end;

Initialization
  registerclasses ( [ TOF_CPTRAITEMENTIAS ] ) ;
end.
