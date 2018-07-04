{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/11/2002
Modifié le ... : 01/09/2004
Description .. : Source TOF de la FICHE : UTOFCPJALECR ()
Suite ........ : - CA - 13/05/2004 : Ajout de $IFDEF COMPTA pour ne
Suite ........ : pas récupérer TofMeth en appel depuis un produit autre
Suite ........ : que la comptabilité.
Suite ........ : - CA - 10/08/2004 : IFDEF COMPTA remplacé par IFDFEF 
Suite ........ : GCGC pour avoir l'héritage de TOFMETH en gestion des 
Suite ........ : standards (FQ 14103)
Suite ........ : - CA - 01/09/2004 : FQ 14449 - Ne pas afficher 
Suite ........ : paramétrage de l'état en mode PCL
Mots clefs ... : TOF;UTOFCPJALECR
*****************************************************************}
Unit UTOFCPJALECR ;

Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    Maineagl,
    eQRS1,
    UTOB,
{$ELSE}
    Fe_main,
    QRS1,
    DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
{$IFNDEF CCADM}
{$IFNDEF GCGC}
    tofmeth,
{$ENDIF}
{$ENDIF}
    Ent1,
    ULibWindows,
    ULibExercice,
    AglInit,      // TheData
    UTOF
    ;

Type
{$IFNDEF CCADM}
{$IFNDEF GCGC}
    TOF_CPJALECR  = Class (TOF_METH)
{$ELSE}
    TOF_CPJALECR  = Class (TOF)
{$ENDIF}
{$ELSE}
    TOF_CPJALECR  = Class (TOF)
{$ENDIF}
   private
    E_EXERCICE          : THValComboBox ;
    E_DATECOMPTABLE     : THEdit ;
    E_DATECOMPTABLE_    : THEdit ;
    E_QUALIFPIECE       : THMultiValComboBox ;
    LIBEQUALIFPIECE     : THEdit ;
    AFFDEVISE           : THEdit ;
    DEVPIVOT            : THEdit ;
    XX_RUPTURE          : THEdit ;
    E_DEVISE            : THValComboBox ;
    TRIDATE             : TCheckBox ;
    DateD, DateF        : TDatetime ;
    AFFICHAGE           : THRadioGroup;
    FMP                 : TCheckBox ;
    procedure DateOnExit     ( Sender : TObject ) ;
    procedure ExoOnChange    ( Sender : TObject ) ;
    procedure DeviseOnChange ( Sender : TObject ) ;
    {$IFNDEF CCADM}
    {$IFNDEF GCGC}
    procedure FTimerTimer    ( Sender : TObject ) ;
    {$ENDIF}
    {$ENDIF}

   public
    procedure OnLoad                              ; override ;
    procedure OnUpdate                            ; override ;
    procedure OnArgument (S : String )            ; override ;
    procedure OnClose                             ; override ;
    procedure OnNew                               ; override ;
    {$IFNDEF CCADM}
    {$IFNDEF GCGC}
    procedure OnChangeFiltre ( Sender : TObject ) ; override ;
    procedure ChargementCritEdt ; override ; // JP 14/10/05 : FQ 16864
    procedure MySelectFiltre    ; override ;
    {$ENDIF}
    {$ENDIF}
  end ;

function CPLanceFiche_CPJALECR(vStParam : string = '' ) : string ;

Implementation

uses
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
{$IFNDEF GCGC}
  CritEdt,
{$ENDIF}
  UtilPGI;

function CPLanceFiche_CPJALECR(vStParam : string = '' ) : string ;
begin
 result := AGLLanceFiche('CP','CPJALECR','','',vStParam);
end;


procedure TOF_CPJALECR.OnLoad ;
var
    lstListeTypeEcriture : string;
    lstTypeEcriture : string;
begin

  Inherited ;


  if VH^.PaysLocalisation=CodeISOes then
     XX_RUPTURE.Text:=''
  else begin
    if TRIDATE.Checked then XX_RUPTURE.Text := ' E_EXERCICE,e_PERIODE,E_JOURNAL,E_NUMEROPIECE,E_DATECOMPTABLE,E_NUMLIGNE '
    else XX_RUPTURE.Text := ' E_EXERCICE,E_PERIODE,E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE ' ;
  End ; //XVI 24/02/2005

  // GCO - 04/02/2003
  // Traduction de la MultiValComboBox pour le type des écritures
  LibEQualifPiece.Text := '';
  if not E_QualifPiece.Tous then
  begin
    lStListeTypeEcriture := E_QualifPiece.Text;
    while Pos(';', lStListeTypeEcriture ) > 0 do
    begin
      lStTypeEcriture        := ReadTokenSt(lStListeTypeEcriture);
      LibEQualifPiece.Text   := LibEQualifPiece.Text + RechDom(E_QualifPiece.DataType, lStTypeEcriture, False) + ', ' ;
    end;
    LibEQualifPiece.Text := Copy(LibEQualifPiece.Text, 0, Length(LibEQualifPiece.Text)-2);
  end
  else
    LibEQualifPiece.Text := '<<' + TraduireMemoire('Tous') + '>>';

  {$IFNDEF CCADM}
  {$IFNDEF GCGC}
  // GCO - 19/09/2006 - Gestion BOI (Norme NF)
  FDateFinEdition := StrToDate(E_DateComptable_.Text);
  {$ENDIF}
  {$ENDIF}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/01/2003
Modifié le ... : 31/01/2003
Description .. : LG - 29/01/2003 - FB 11870 - utilisation de l'exercice de ref
Suite ........ : en mode pcl
Suite ........ :
Suite ........ : LG - 31/01/2003 - FB 11857 - correction du libelle du
Suite ........ : caption de tri par date.
Suite ........ : - test si l'exercice de ref est definie
Suite ........ :
Suite ........ : YM 23/08/2005 FQ 15618 : on ne prend pas les journaux analytiques
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJALECR .OnArgument (S : String ) ;
var
    //QListeEtab : TQuery;
    stToken : string;
    stParam, stValeur : string;
    stWhereJal : string;
    stWhereJalsansAna : String;
begin

  Inherited ;

  // YM 23/08/2005 FQ 15618 : on ne prend pas les journaux analytiques
  stWhereJalsansAna := 'J_NATUREJAL <> "ODA" AND J_NATUREJAL <> "ANA"';

  { Récupération des paramètres }
  { Paramètres à passer sous la forme :
      ex : NATUREJAL=ACH,OD; ==> seuls les journaux de type Achat et OD seront proposés dans la combo E_JOURNAL }
  stWhereJal := '';
  while S<>'' do
  begin
    stToken := ReadTokenSt(S);
    stParam := ReadTokenPipe(stToken,'=');
    if stParam='NATUREJAL' then
    begin
      while ( stToken <> '' ) do
      begin
        stValeur := ReadTokenPipe(stToken, ',');
        if stValeur <> '' then
        begin
          if stWhereJal <> '' then stWhereJal := stWhereJal + ' OR ';
          stWhereJal := stWhereJal+' J_NATUREJAL="'+stValeur+'" ';
        end;
      end;
    end;
  end;

  if stWhereJal = '' then
      stWhereJal := stWhereJalsansAna
  else
      stWhereJal := stWhereJalsansAna + ' AND (' + stWhereJal + ')';

  Ecran.HelpContext        := 7394000;       // BPY 07/02/2003

  E_DATECOMPTABLE          := THEdit(GetControl('E_DATECOMPTABLE')) ;
  E_DATECOMPTABLE_         := THEdit(GetControl('E_DATECOMPTABLE_')) ;
  E_EXERCICE               := THValComboBox(GetControl('E_EXERCICE')) ;
  CInitComboExercice(E_EXERCICE);
  E_DEVISE                 := THValComboBox(GetControl('E_DEVISE')) ;
  E_QUALIFPIECE            := THMultiValComboBox(GetControl('E_QUALIFPIECE'));
  AFFICHAGE	 	   := THRadioGroup(GetControl('AFFICHAGE'));
  XX_RUPTURE               := THEdit(GetControl('XX_ORDERBY')) ;
  DEVPIVOT                 := THEdit(GetControl('DEVPIVOT')) ;
  LIBEQUALIFPIECE          := THEdit(GetControl('LIBEQUALIFPIECE'));
  TRIDATE                  := TCheckBox(GetControl('TRIDATE')) ;
  AFFDEVISE                := THEdit(GetControl('AFFDEVISE')) ;
  FMP                      := TCheckBox(GetControl('MP',true));

  TRIDATE.visible:=(VH^.PaysLocalisation<>CodeISOes) ; //XVI 24/02/2005

  NotifyErrorComponent(E_DATECOMPTABLE    , 'E_DATECOMPTABLE') ;
  NotifyErrorComponent(E_DATECOMPTABLE_   , 'E_DATECOMPTABLE_') ;
  NotifyErrorComponent(E_EXERCICE         , 'E_EXERCICE') ;
  NotifyErrorComponent(E_DEVISE           , 'E_DEVISE') ;
  NotifyErrorComponent(E_QUALIFPIECE      , 'E_QUALIFPIECE' );
  NotifyErrorComponent(AFFICHAGE          , 'AFFICHAGE') ;
  NotifyErrorComponent(AFFDEVISE          , 'AFFDEVISE') ;
  NotifyErrorComponent(DEVPIVOT           , 'DEVPIVOT') ;
  NotifyErrorComponent(LIBEQUALIFPIECE    , 'LIBEQUALIFPIECE') ;

  E_EXERCICE.OnChange      := ExoOnChange ;
  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
   E_EXERCICE.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
    else
     E_EXERCICE.Value      := CExerciceVersRelatif(VH^.Entree.Code) ;
  E_DATECOMPTABLE.OnExit   := DateOnExit ;
  E_DATECOMPTABLE.OnExit   := DateOnExit ;
  E_DEVISE.OnChange        := DeviseOnChange ;

  SetControlProperty ('E_JOURNAL','Plus',stWhereJal);
  if VH^.PaysLocalisation=CodeISOes then
     THMultiValcomboBox(GetControl('E_JOURNAL')).Value:='' ; //XVI 24/02/2005

  // Etats Chainés
  {$IFNDEF CCADM}
  {$IFNDEF GCGC}
  FTimer.OnTimer           := FTimerTimer ;
  {$ENDIF}
  {$ENDIF}
  FMP.Checked              := not ( CtxPCl in V_PGI.PgiContexte ) ;


  // CA - 01/09/2004 - FQ 14449
  SetControlVisible('FEtat',not (ctxStandard in V_PGI.PGIContexte));
  SetControlVisible('TEtat',not (ctxStandard in V_PGI.PGIContexte));
  SetControlVisible('BParamEtat',not (ctxStandard in V_PGI.PGIContexte) and (not (ctxPCL in V_PGI.PGIContexte)));
  if (ctxaffaire in V_PGI.Pgicontexte) or  (ctxgcaff in V_PGI.Pgicontexte) then
    begin        //mcd 10/10/2006 equalite GIGA13418  + FQ11715  le 29/06/07
    Ecran.Caption := TraduireMemoire('Journal de ventes comptable');
    updatecaption(Ecran);
    end;
end ;

procedure TOF_CPJALECR .OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPJALECR.DateOnExit(Sender: TObject);
begin
 {$IFNDEF CCADM}
 {$IFNDEF GCGC}
 DoDateOnExit(THEdit(Sender), E_DATECOMPTABLE, E_DATECOMPTABLE_, DateD, DateF);
 {$ENDIF}
 {$ENDIF}
end;

procedure TOF_CPJALECR.ExoOnChange(Sender: TObject);
begin
  if ((E_EXERCICE<>nil) and (E_DATECOMPTABLE<>nil) and (E_DATECOMPTABLE_<>nil)) then
    CExoRelatifToDates ( E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  DateD := StrToDate(E_DATECOMPTABLE.Text);
  DateF := StrToDate(E_DATECOMPTABLE_.Text);
end;

procedure TOF_CPJALECR.DeviseOnChange(Sender: TObject);
begin
 AFFICHAGE.Enabled := ( E_DEVISE.Value <> '' ) and ( E_DEVISE.Value <> V_PGI.DevisePivot );
end;

procedure TOF_CPJALECR.OnUpdate;  
begin
  inherited;
  // Affichage en DEVISE ou en DEVISEPIVOT
  AFFICHAGE.Items[0] := '&' + VH^.LibDevisePivot;
  if (AFFICHAGE.Enabled) and (AFFICHAGE.ItemIndex=1) then
  begin  // affichage en devise
    DEVPIVOT.Text  := TraduireMemoire('Devise'); // GCO - 04/02/2003 FB : 11876
    AFFDEVISE.Text := 'X'
  end
  else
  begin
    DEVPIVOT.Text  := RechDom('TTDEVISE',V_PGI.DevisePivot,false) ;
    AFFDEVISE.Text := '-' ;
  end;

  TFQRS1(Ecran).WhereSQL :=CMajRequeteExercice (  E_EXERCICE.Value,TFQRS1(Ecran).WhereSQL);

  {$IFNDEF CCADM}
  {$IFNDEF GCGC}
  //06/12/2006 YMO Norme NF
  if FProvisoire.Text='' then
  {$ENDIF}
  {$ENDIF}
{$IFNDEF CERTIFNF}
     CPEnregistreLog('JALECR '+E_DATECOMPTABLE.Text+' - '+E_DATECOMPTABLE_.Text);
{$ELSE}
     CPEnregistreJalEvent('CEE','Edition journal des écritures','JALECR '+E_DATECOMPTABLE.Text+' - '+E_DATECOMPTABLE_.Text);
{$ENDIF}


end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/02/2003
Modifié le ... :   /  /    
Description .. : Correction des initialisations des combos
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJALECR.OnNew;
begin
  inherited;

  // ----------------- Init des composants si pas de Filtre DEFAUT -------------
  {$IFNDEF CCADM}
  {$IFNDEF GCGC}
  if FFiltres.Text = '' then
  begin
    E_QualifPiece.Text        := 'N;';
    {JP 04/10/05 : FQ 16149 : si l'établissement est positionné (PositionneEtabUser),
                   il est préférable d'éviter de le "dépositionner" }
    if (ComboEtab.Enabled) and (ComboEtab.ItemIndex < 0) then
      ComboEtab.ItemIndex := 0;

    E_Devise.ItemIndex        := 0;
  end;
  {$ENDIF}
  {$ENDIF}
  
  DeviseOnChange(nil);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/02/2003
Modifié le ... : 06/09/2005
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFNDEF CCADM}
{$IFNDEF GCGC}
procedure TOF_CPJALECR.FTimerTimer(Sender: TObject);
begin
  if FCritEdtChaine <> nil then
  begin
    with FCritEdtChaine do
    begin
      if CritEdtChaine.UtiliseCritStd then
      begin
        E_Exercice.Value      := CritEdtChaine.Exercice.Code;
        E_DateComptable.Text  := DateToStr(CritEdtChaine.Exercice.Deb);
        E_DateComptable_.Text := DateToStr(CritEdtChaine.Exercice.Fin);

        // GCO - 03/08/2007 - FQ 19377 
        THMultiValcomboBox(GetControl('E_QUALIFPIECE')).Value := CritEdtChaine.TypeEcriture;
      end;
    end;
  end;
  inherited;
end;

procedure TOF_CPJALECR.OnChangeFiltre(Sender: TObject);
begin
  inherited;
  //CExoRelatifToDates ( E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_, True);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFNDEF CCADM}
{$IFNDEF GCGC}
procedure TOF_CPJALECR.ChargementCritEdt;
begin
  inherited;
  // Récupération des critères d'impression
  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    SetControlText('E_JOURNAL', ClassCritEdt(TheData).CritEdt.Cpt1);
    SetControlText('E_EXERCICE', CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code));

    if ClassCritEdt(TheData).CritEdt.Jal.NumPiece1 >0 then
      SetControlText('E_NUMEROPIECE' , IntToStr(ClassCritEdt(TheData).CritEdt.Jal.NumPiece1));

    if ClassCritEdt(TheData).CritEdt.Jal.NumPiece2 > 0 then
      SetControlText('E_NUMEROPIECE_', IntToStr(ClassCritEdt(TheData).CritEdt.Jal.NumPiece2));

    E_DATECOMPTABLE.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    E_DATECOMPTABLE_.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);

    E_QUALIFPIECE.Text := ClassCritEdt(TheData).CritEdt.QualifPiece;
    ComboEtab.Value := ClassCritEdt(TheData).CritEdt.Etab;
    E_DEVISE.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;
    TheData := nil;
  end;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... : 20/01/2006
Description .. :
Suite ........ : FQ 17314 - Obligé de créer un MySelecFiltre suite à des
Suite ........ : modifications dans les sources de l'AGL. En effet, l'appel
Suite ........ : à CExoRelatifToDates ne suffit plus.
Mots clefs ... :
*****************************************************************}
{$IFNDEF CCADM}
{$IFNDEF GCGC}
procedure TOF_CPJALECR.MySelectFiltre;
var lExoDate : TExoDate;
begin
  inherited;
  lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));

  if (CtxExercice.QuelExo(E_DateComptable.Text, False) <> lExoDate.Code) or
     (CtxExercice.QuelExo(E_DateComptable_.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  end;
end;
{$ENDIF}
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPJALECR  ] ) ;
end.


