{***********UNITE*************************************************
Auteur  ...... : GF
Créé le ...... : 15/11/2000
Modifié le ... : 19/02/2001
Description .. : Source TOM de la TABLE : BTETAT (BTETAT)
Mots clefs ... : TOM;BTETAT
*****************************************************************}
Unit UTOMBTETAT;

Interface

Uses Forms,
     HeureUtil,
     SysUtils,
     Classes,
     Controls,
     graphics,
     StdCtrls,
     ExtCtrls,
     MsgUtil,
     HCtrls,
     UTOM,
     UTOB,
     HTB97,
     {$IFDEF EAGLCLIENT}
     MaineAGL,
     eFiche,
     eFichList,
     UtileAGL,
     {$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     Fe_main,
     {$ENDIF}
     HRListeIcone;

Type
  TOM_BTETAT = Class (TOM)
    procedure OnArgument ( S: String )   ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  private

  	NomForm	: String;

    //Gestion des boutons
    BFOND	 	  : TToolbarButton97;
    BCOULEUR	: TToolbarButton97;
    BIMGICONE	: TToolbarButton97;
    BFONTE		: TToolbarButton97;
    IMGICONE	: TImage;

    LFOND			: THEdit;
    LCOULEUR	: THEdit;
    DUREE   	: THEdit;

    LFONTE		: THLabel;

    procedure BFond_OnClick(Sender: TObject);
    procedure BFonte_OnClick(Sender: TObject);
    procedure BCouleur_OnClick(Sender: TObject);
	  procedure BImgIcone_OnClick(Sender: TObject);
    procedure Duree_OnExit(Sender: TObject);
    function EventExists (CodeEtat : string) : boolean;
  end ;

Implementation

procedure TOM_BTETAT.OnArgument ( S: String ) ;
begin
  Inherited ;

  NomForm 	:= ecran.Name;

  BFond 		:= TToolbarButton97(ecran.FindComponent('B_FOND'));
  BFond.onclick := BFond_OnClick;

  BCouleur 	:= TToolbarButton97(ecran.FindComponent('B_COULEUR'));
  BCouleur.onclick := BCouleur_OnClick;

  BFonte		:= TToolbarButton97(ecran.FindComponent('BFONTE'));
  BFonte.onclick := BFonte_OnClick;

  BImgIcone := TToolbarButton97(ecran.FindComponent('B_ICONE'));
  BImgIcone.onclick := BImgIcone_OnClick;

  ImgIcone  := Timage(ecran.FindComponent('IMAGEICONE'));

  LFond 		:= THEdit(ecran.FindComponent('LFOND'));
  LCouleur 	:= THEdit(ecran.FindComponent('LCOULEUR'));

  DUREE	  	:= THEdit(ecran.FindComponent('DUREE'));
  DUREE.OnChange := DUREE_OnExit;

	LFonte 		:= THLabel(ecran.FindComponent('LFONTE'));
end ;

procedure TOM_BTETAT.OnNewRecord ;
begin
  Inherited ;

  // Init. couleurs ,fontes,icone
  Setfield('BTA_NUMEROICONE',9999);
  Setfield('BTA_COULEURFOND',Inttostr(16777215));
  Setfield('BTA_COULEUR',IntToStr(0));

  AfficheCouleur(LFOND, THEdit(GetControl('BTA_COULEURFOND')));
  AfficheCouleur(LCOULEUR, THEdit(GetControl('BTA_COULEUR')));

  SetField('BTA_FONTE','Ms Sans Serif');
  SetField('BTA_FONTESIZE',8);
  SetField('BTA_FONTESTYLE','fsBold');

  LFonte.Font.Name := 'Ms Sans Serif';
  LFonte.Font.Size := 8;
  LFonte.Font.Style := [fsBold];
  LFonte.Font.Color := LCouleur.Color;

  SetField('BTA_DUREEMINI',1);
  Duree.Text := '01:00';
  SetField('BTA_ASSOSRES','X');
end ;

procedure TOM_BTETAT.OnDeleteRecord ;
begin
  Inherited ;

  //Ajouter ici la lecture en fonction des évènements
  If ExisteSQL('Select BEP_BTETAT from BTEVENPLAN where BEP_BTETAT="'+GetField('BTA_BTETAT')+'"') then
	   begin
     AfficheErreur(NomForm, '4', 'Fiche type d''action');
     exit;
     end;

end ;

procedure TOM_BTETAT.OnUpdateRecord ;
Var Temp 			: Double;
    DureeMini	: TDateTime;
begin
  Inherited ;

  //Mise à jour de la durée mini
  DureeMini := StrToTime(Duree.Text);
  Temp := TimeToFloat(DureeMini, true);
  SetField('BTA_DUREEMINI', Temp);
  // Fonte & Couleur
  SetField('BTA_COULEURFOND', LFOND.Color);
  SetField('BTA_COULEUR', LCOULEUR.Color);

  SetField('BTA_FONTE',LFONTE.Font.Name);
  SetField('BTA_FONTESIZE',LFONTE.Font.Size);

  If fsBold in LFONTE.Font.Style Then
     SetField('BTA_FONTESTYLE','fsBold')
  else if fsItalic in LFONTE.Font.Style Then
     SetField('BTA_FONTESTYLE','fsItalic')
  else if fsUnderline in LFONTE.Font.Style Then
     SetField('BTA_FONTESTYLE','fsUnderline')
  else if fsStrikeout in LFONTE.Font.Style Then
     SetField('BTA_FONTESTYLE','fsStrikeout')
  Else
     SetField('BTA_FONTESTYLE','');

end;

procedure TOM_BTETAT.OnLoadRecord ;
var ImageIcone	: TImage;
    sttmp 			: String;
		Temp 				: Double;
    DureeMini		: TDateTime;
begin
  Inherited ;

  //Chargement de la durée mini
  Temp := GetField('BTA_DUREEMINI');
  DureeMini := FloatToTime(Temp, true);
  Duree.text := TimeToStr(DureeMini);

  // Chargement de l'icône associé
  ImageIcone:=TImage(TFfiche(Ecran).FindComponent('IMAGEICONE'));
  If GetField('BTA_NUMEROICONE') <> 9999 then
     ChargeIcone(GetField('BTA_NUMEROICONE'),ImageIcone)
  else
     ImageIcone.Picture:= Nil;

  // Chargement Couleurs & Fonte
  AfficheCouleur(LFOND, THEdit(GetControl('BTA_COULEURFOND')));
  AfficheCouleur(LCOULEUR,THEdit(GetControl('BTA_COULEUR')));

  LFONTE.Font.Name := GetField('BTA_FONTE');
  LFONTE.Font.Color := LCOULEUR.Color;
  StTmp := GetField('BTA_FONTESTYLE');

  if StTmp <> '' Then
     begin
     If StTmp = 'fsBold' Then
        LFONTE.Font.Style := [fsBold]
     else if StTmp = 'fsItalic' Then
        LFONTE.Font.Style := [fsItalic]
     else if StTmp = 'fsUnderline' Then
        LFONTE.Font.Style := [fsUnderline]
     else if StTmp = 'fsStrikeout' Then
        LFONTE.Font.Style := [fsStrikeout];
     End;

  StTmp := GetField('BTA_FONTESIZE');
  if (StrToInt (StTmp) >= 1) Then LFONTE.Font.Size := StrToInt(StTmp);
  if (GetField('BTA_BTETAT')<>'ABSPAIE') and (GetField('BTA_BTETAT')<>'ACT-GRC') and (not (EventExists (GetField('BTA_BTETAT'))))  then
  	TToolbarButton97(GetControl('Bdelete')).enabled := true;
end ;

procedure TOM_BTETAT.BFOND_OnClick(Sender: TObject);
begin

  SelColorNew(LFOND,THEdit(GetControl('BTA_COULEURFOND')), TForm(Ecran));
  if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;

end;

procedure TOM_BTETAT.BCOULEUR_OnClick(Sender: TObject);
begin

  SelColorNew(LCOULEUR,THEdit(GetControl('BTA_COULEUR')), TForm(Ecran));

  //Affichage de la police dans la nouvelle couleur
  LFONTE.Font.Color := LCOULEUR.Color;
  if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;

end;

procedure TOM_BTETAT.BFONTE_OnClick(Sender: TObject);
begin

	SelFonteNew(LFONTE, THEdit(GetControl('BTA_FONTE')), TForm(Ecran));

  //Affichage de la couleur en fonction de la police
  LCOULEUR.Color := LFONTE.Font.Color;
  if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;

end;

procedure TOM_BTETAT.BImgIcone_OnClick(Sender: TObject);
var NumIcone : Variant;
begin

  ImgIcone.Picture := nil;

  NumIcone:=ListeIcone(True);

  If (NumIcone <> -1) or (NumIcone=9999) then ChargeIcone(NumIcone,ImgIcone);

	If NumIcone <> -1 then
  begin
  	if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;
  	SetField('BTA_NUMEROICONE',NumIcone);
  end;

end;

procedure TOM_BTETAT.Duree_OnExit(Sender: TObject);
Begin
	if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;
end;

procedure TOM_BTETAT.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_BTETAT.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_BTETAT.OnClose;
begin
  inherited;
end;

function TOM_BTETAT.EventExists(CodeEtat: string): boolean;
begin
  result := ExisteSQL ('SELECT BEP_CODEEVENT FROM BTEVENPLAN WHERE BEP_BTETAT="'+CodeEtat+'"');
end;

Initialization
  registerclasses ( [ TOM_BTETAT ] ) ;
end.
