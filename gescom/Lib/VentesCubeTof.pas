{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCVENTES ()
Mots clefs ... : TOF;GCVENTES
*****************************************************************}
Unit VentesCubeTof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     sysutils,
     ComCtrls,
     Forms,
{$IFDEF EAGLCLIENT}
      Utob,
      eMul,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      mul,
{$ENDIF}
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     HTB97,
     paramSoc,
     UtilGC;

Type
  TOF_GCVENTES = Class (TOF)

    procedure OnArgument (S : String ) ; override ;

  Private

    Affaire   : THEdit;
    Affaire0  : THEdit;
    Affaire1  : THEdit;
    Affaire2  : THEdit;
    Affaire3  : THEdit;
    Avenant   : THEdit;
    Client    : THEdit;
    //
    EtatAffaire : THValComboBox;
    StWhereEtat : string;
    //
    BTSelect  : TToolbarbutton97;
    BtEfface  : TToolbarButton97;
    //
    procedure BtEffaceOnclick(Sender: TObject);
    procedure BtSelectOnClick(Sender: TObject);
    procedure EtatAffaireOnChange(Sender: TObject);
    procedure BRechResponsable(Sender: TObject);
    //

    end ;

Implementation

uses affaireutil,
     FactUtil,
     Utilressource;

procedure TOF_GCVENTES.OnArgument (S : String ) ;
Var i         : integer;
    AuMoinsUn : boolean;
    SPlus     : string;
    CC        : THValComboBox;
    CCE : Thedit;
begin
  Inherited ;

  {for i:=1 to 10 do
    BEGIN
    if i<10 then Suf:=IntToStr(i) else Suf:='A' ;
    Nam:='GP_LIBRETIERS'+Suf ; GCTitreZoneLibre(Nam,Titre) ; SetControlCaption('T'+Nam,Titre);
    Nam:='GL_LIBREART'+Suf   ; GCTitreZoneLibre(Nam,Titre) ; SetControlCaption('T'+Nam,Titre);
    END ;}

  AuMoinsUn:=False;

  For i:=1 to 10 do if i<10 then
    AuMoinsUn:=(ChangeLibre2('TGP_LIBRETIERS'+intToStr(i),TForm(Ecran)) or AuMoinsUn)
  else
    AuMoinsUn:=(ChangeLibre2('TGP_LIBRETIERSA',TForm(Ecran))or AuMoinsUn);

//uniquement en line
//
{*
	setcontrolproperty('GL_REPRESENTANT', 'Visible', False);
	setcontrolproperty('GL_ETABLISSEMENT', 'Visible', False);


 	setcontrolproperty('GP_LIBRETIERS1', 'Visible', False);
  setcontrolproperty('GP_LIBRETIERS2', 'Visible', False);
  setcontrolproperty('GP_LIBRETIERS3', 'Visible', False);
  setcontrolproperty('GP_LIBRETIERS4', 'Visible', False);
  setcontrolproperty('GP_LIBRETIERS5', 'Visible', False);
//
 	setcontrolproperty('GL_LIBREART1', 'Visible', False);
  setcontrolproperty('GL_LIBREART2', 'Visible', False);
  setcontrolproperty('GL_LIBREART3', 'Visible', False);
  setcontrolproperty('GL_LIBREART4', 'Visible', False);
  setcontrolproperty('GL_LIBREART5', 'Visible', False);
//
	setcontrolproperty('TGL_REPRESENTANT', 'Visible', False);
	setcontrolproperty('TGL_ETABLISSEMENT', 'Visible', False);
//
 	setcontrolproperty('TGP_LIBRETIERS1', 'Visible', False);
  setcontrolproperty('TGP_LIBRETIERS2', 'Visible', False);
  setcontrolproperty('TGP_LIBRETIERS3', 'Visible', False);
  setcontrolproperty('TGP_LIBRETIERS4', 'Visible', False);
  setcontrolproperty('TGP_LIBRETIERS5', 'Visible', False);
//
 	setcontrolproperty('TGL_LIBREART1', 'Visible', False);
  setcontrolproperty('TGL_LIBREART2', 'Visible', False);
  setcontrolproperty('TGL_LIBREART3', 'Visible', False);
  setcontrolproperty('TGL_LIBREART4', 'Visible', False);
  setcontrolproperty('TGL_LIBREART5', 'Visible', False);

*}

  Affaire       := THEdit(GetControl('GL_AFFAIRE'));
  Affaire0      := THEdit(GetControl('AFFAIRE0'));
  Affaire1      := THEdit(GetControl('GL_AFFAIRE1'));
  Affaire2      := THEdit(GetControl('GL_AFFAIRE2'));
  Affaire3      := THEdit(GetControl('GL_AFFAIRE3'));
  Avenant       := THEdit(GetControl('GL_AVENANT'));
  Client        := THEdit(GetCOntrol('GL_TIERS'));

  EtatAffaire   := THValComboBox(Getcontrol('ETATAFFAIRE'));
  EtatAffaire.OnChange := EtatAffaireOnChange;

  BTSelect      := TToolbarbutton97(GetControl('BSELECT1'));
  BtSelect.OnClick     := BtSelectOnClick;

  BtEfface      := TToolbarbutton97(GetControl('BEFFACE1'));
  BtEfface.OnClick     := BtEffaceOnclick;

  MajChampsLibresArticle(TForm(Ecran),'GL',(not AuMoinsUn),'PCOMPLEMENTS','PCOMPLEMENTS');


  CC:=THValComboBox(GetControl('GL_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

  CC:=THValComboBox(GetControl('GL_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

{$IFDEF BTP}
  If (S = 'I') or (S = 'W') then      //contrat ou intervention
  begin
    THEdit(GetControl('GL_AFFAIRE')).Plus := 'AND GL_AFFAIRE LIKE "' + S + '%"';
    THMultiValComboBox(GetControl('GL_NATUREPIECEG')).Plus := 'AND (GPP_NATUREPIECEG="FAC" OR GPP_NATUREPIECEG="AVC")';
    SetControlText('XX_WHERE',' GP_NATUREPIECEG IN ("FAC","AVC")');
    //
    sPlus := sPlus + ' AND (ISNUMERIC(CC_LIBRE)=1)';
    Affaire0.text := 'W';
    SetControlText ('TGL_AFFAIRE', 'Code Appel') ;
    ETATAFFAIRE.Plus := sPlus;
    Affaire.Text := '';
  end else                            //Chantier ou autres...
  begin
    THEdit(GetControl('GL_AFFAIRE')).Plus := 'AND (GL_AFFAIRE LIKE "A%" OR GL_AFFAIRE="")';
    //FV1 : 22/05/2015 -  Ajout de la nature DAP (devis interv)
    //THMultiValComboBox(GetControl('GL_NATUREPIECEG')).Plus := 'AND (GPP_NATUREPIECEG="FBT" OR GPP_NATUREPIECEG="ABT" OR GPP_NATUREPIECEG="FAC" OR GPP_NATUREPIECEG="AVC" OR GPP_NATUREPIECEG="DBT" OR GPP_NATUREPIECEG="ABC" OR GPP_NATUREPIECEG="FBC")';
    THMultiValComboBox(GetControl('GL_NATUREPIECEG')).Plus := 'AND (GPP_NATUREPIECEG="FBT" OR GPP_NATUREPIECEG="ABT" OR GPP_NATUREPIECEG="FAC" OR GPP_NATUREPIECEG="AVC" OR GPP_NATUREPIECEG="DBT" OR GPP_NATUREPIECEG="ABC" OR GPP_NATUREPIECEG="FBC" OR GPP_NATUREPIECEG="DAP" )';
    SetControlText('XX_WHERE',' GP_NATUREPIECEG IN ("FBT","ABT","FAC","AVC","DBT","ABC","FBC","DAP")');
    //
    sPlus := sPlus + ' AND (CC_LIBRE="BTP" AND CC_CODE <> "ATT")';
    Affaire0.Text := 'A';
    SetControlText ('TGL_AFFAIRE','Code Chantier') ;
    EtatAffaire.Plus := sPlus;
    Affaire.Text := '';
  end;

  THEdit(GetControl('GL_TIERSFACTURE')).Plus := 'AND T_NATUREAUXI="CLI"';

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, TaModif, Affaire.Text, True);
{$ENDIF}

	CCE:=ThEdit(GetControl('AFF_RESPONSABLE')) ;
	if CCE<>Nil then
  begin
  	PositionneResponsableUser(CCE) ;
    CCE.OnElipsisClick := BRechResponsable;
	end;

end ;

procedure TOF_GCVENTES.BtSelectOnClick(Sender: TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(AFFAIRE0, Affaire1, Affaire2, Affaire3, AVENANT, Client, StChamps, false, false, false, False, False,'') then AFFAIRE.text := Stchamps;

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, TaModif, Affaire.Text, True);

end;

procedure TOF_GCVENTES.BtEffaceOnclick(Sender: TObject);
begin

  Affaire0.text := '';
  Affaire1.Text := '';
  Affaire2.text := '';
  Affaire3.text := '';
  Avenant.text  := '';
  Affaire.text  := '';

end;

procedure TOF_GCVENTES.EtatAffaireOnChange(Sender: TObject);
var StWhere : string;
    st      : string;
    P       : Integer;
begin

  StWhere := GetControlText('XX_WHERE');

  //on vérifie si l'état n'est pas déjà dans la clause where...
  P:= pos('ETATAFFAIRE',Stwhere);
  if P<> 0  then
  begin
    St := Copy(Stwhere,P, Length(StWhere)-P);
    SetLength(Stwhere, Length(StWhere)- (Length(St)+ 5));
  end;
  if EtatAffaire.Value <> '' then
  begin
    if StWhere <> '' then StWhere := StWhere + ' AND';

    StWhere := StWhere + ' ETATAFFAIRE="' + EtatAffaire.Value + '"';
    SetControlText('XX_WHERE', StWhere);
  end;

end;


procedure TOF_GCVENTES.BRechResponsable(Sender: TObject);
Var QQ  : TQuery;
    SS  : THCritMaskEdit;
begin

  if GetParamSocSecur('SO_AFRECHRESAV', True) then
  begin
    SS := THCritMaskEdit.Create(nil);
    GetRessourceRecherche(SS,'ARS_RESSOURCE=' + ThEdit(Sender).text + ';TYPERESSOURCE=SAL', '', '');
    if (SS.Text <> THEdit(Sender).text) then
    begin
      if SS.text = '' then ss.text := ThEdit(Sender).text;
    end;
    ThEdit(Sender).text  := SS.Text;
    SS.Free;
  end
  else
    GetRessourceRecherche(ThEdit(Sender),'ARS_TYPERESSOURCE="SAL"', '', '');

end;


Initialization
  registerclasses ( [ TOF_GCVENTES ] ) ;
end.

