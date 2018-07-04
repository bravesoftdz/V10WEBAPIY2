{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/03/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAFFDOC ()
Mots clefs ... : TOF;BTPARAFFDOC
*****************************************************************}
Unit BTPARAFFDOC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UFactgestionAff,
     HTB97,
     BtpUtil,
     UTOF ; 

Type
  TOF_BTPARAFFDOC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TheModele : TAffichageDoc;
    TheCourant : TAffichageDoc;
    procedure AffecteNormalToScreen;
    procedure SetEvents;
    procedure Brushniv1Click (Sender : Tobject);
    procedure Fontniv1Click(Sender: TObject);
    procedure Colorniv1Click(Sender: TObject);
    procedure Brushniv2Click(Sender: Tobject);
    procedure Brushniv3Click(Sender: Tobject);
    procedure Brushniv4Click(Sender: Tobject);
    procedure Brushniv5Click(Sender: Tobject);
    procedure Brushniv6Click(Sender: Tobject);
    procedure Brushniv7Click(Sender: Tobject);
    procedure Brushniv8Click(Sender: Tobject);
    procedure Brushniv9Click(Sender: Tobject);
    procedure Colorniv2Click(Sender: TObject);
    procedure Colorniv3Click(Sender: TObject);
    procedure Colorniv4Click(Sender: TObject);
    procedure Colorniv5Click(Sender: TObject);
    procedure Colorniv6Click(Sender: TObject);
    procedure Colorniv7Click(Sender: TObject);
    procedure Colorniv8Click(Sender: TObject);
    procedure Colorniv9Click(Sender: TObject);
    procedure Fontniv2Click(Sender: TObject);
    procedure Fontniv3Click(Sender: TObject);
    procedure Fontniv4Click(Sender: TObject);
    procedure Fontniv5Click(Sender: TObject);
    procedure Fontniv6Click(Sender: TObject);
    procedure Fontniv7Click(Sender: TObject);
    procedure Fontniv8Click(Sender: TObject);
    procedure Fontniv9Click(Sender: TObject);
    procedure BrushDetOuvClick(Sender: Tobject);
    procedure BrushOuvClick(Sender: Tobject);
    procedure BrushVarClick(Sender: Tobject);
    procedure ColorDetOuvClick(Sender: TObject);
    procedure ColorOuvClick(Sender: TObject);
    procedure ColorVarClick(Sender: TObject);
    procedure FontDetOuvClick(Sender: TObject);
    procedure FontOuvClick(Sender: TObject);
    procedure FontVarClick(Sender: TObject);
    procedure DESDETAILChange (Sender : Tobject);
    procedure DEPARAGChange (Sender : Tobject);
    procedure SetScreenToNormal;
    procedure BrushLigClick(Sender: Tobject);
    procedure FontLigClick(Sender: TObject);
    procedure ColorLigClick(Sender: TObject);
    procedure BrushComClick(Sender: Tobject);
    procedure FontComClick(Sender: TObject);
    procedure ColorComClick(Sender: TObject);
    procedure BrushStotClick(Sender: Tobject);
    procedure FontStotClick(Sender: TObject);
    procedure ColorSTotClick(Sender: TObject);

  end ;

Implementation

uses Graphics, Spin;

procedure TOF_BTPARAFFDOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAFFDOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAFFDOC.OnUpdate ;
begin
  Inherited ;
  SetScreenToNormal;
end ;

procedure TOF_BTPARAFFDOC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAFFDOC.OnArgument (S : String ) ;
begin
  Inherited ;
  TheModele :=TAffichageDoc.create (ecran);
  TheModele.gestion := TtaDefaut;
  //
  TheCourant := TAffichageDoc.create (ecran);
  TheCourant.gestion := TtaNormal;
  AffecteNormalToScreen;
	SetEvents;
end ;

procedure TOF_BTPARAFFDOC.OnClose ;
begin
	TheModele.free;
  TheCourant.free;
  Inherited ;
end ;

procedure TOF_BTPARAFFDOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAFFDOC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAFFDOC.AffecteNormalToScreen;
begin
  //
  THEdit(getControl('TPARNIV1')).Color :=  TheCourant.Paragraphe1.brush;
  THEdit(getControl('TPARNIV1')).Font.Name  :=  TheCourant.Paragraphe1.Font.Name;
  THEdit(getControl('TPARNIV1')).Font.Size  :=  TheCourant.Paragraphe1.Font.Size;
  THEdit(getControl('TPARNIV1')).Font.Style  :=  TheCourant.Paragraphe1.Font.Style;
  THEdit(getControl('TPARNIV1')).Font.Color  :=  TheCourant.Paragraphe1.Font.Color;
  //
  THEdit(getControl('TPARNIV2')).Color :=  TheCourant.Paragraphe2.brush;
  THEdit(getControl('TPARNIV2')).Font.Name  :=  TheCourant.Paragraphe2.Font.Name;
  THEdit(getControl('TPARNIV2')).Font.Size  :=  TheCourant.Paragraphe2.Font.Size;
  THEdit(getControl('TPARNIV2')).Font.Style  :=  TheCourant.Paragraphe2.Font.Style;
  THEdit(getControl('TPARNIV2')).Font.Color  :=  TheCourant.Paragraphe2.Font.Color;
  //
  THEdit(getControl('TPARNIV3')).Color :=  TheCourant.Paragraphe3.brush;
  THEdit(getControl('TPARNIV3')).Font.Name  :=  TheCourant.Paragraphe3.Font.Name;
  THEdit(getControl('TPARNIV3')).Font.Style  :=  TheCourant.Paragraphe3.Font.Style;
  THEdit(getControl('TPARNIV3')).Font.size  :=  TheCourant.Paragraphe3.Font.size;
  THEdit(getControl('TPARNIV3')).Font.color  :=  TheCourant.Paragraphe3.Font.color;
  //
  THEdit(getControl('TPARNIV4')).Color :=  TheCourant.Paragraphe4.brush;
  THEdit(getControl('TPARNIV4')).Font.Name  :=  TheCourant.Paragraphe4.Font.Name;
  THEdit(getControl('TPARNIV4')).Font.Style  :=  TheCourant.Paragraphe4.Font.Style;
  THEdit(getControl('TPARNIV4')).Font.size  :=  TheCourant.Paragraphe4.Font.size;
  THEdit(getControl('TPARNIV4')).Font.color  :=  TheCourant.Paragraphe4.Font.color;
  //
  THEdit(getControl('TPARNIV5')).Color :=  TheCourant.Paragraphe5.brush;
  THEdit(getControl('TPARNIV5')).Font.Name  :=  TheCourant.Paragraphe5.Font.Name;
  THEdit(getControl('TPARNIV5')).Font.Style  :=  TheCourant.Paragraphe5.Font.Style;
  THEdit(getControl('TPARNIV5')).Font.size  :=  TheCourant.Paragraphe5.Font.size;
  THEdit(getControl('TPARNIV5')).Font.color  :=  TheCourant.Paragraphe5.Font.color;
  //
  THEdit(getControl('TPARNIV6')).Color :=  TheCourant.Paragraphe6.brush;
  THEdit(getControl('TPARNIV6')).Font.Name  :=  TheCourant.Paragraphe6.Font.Name;
  THEdit(getControl('TPARNIV6')).Font.Style  :=  TheCourant.Paragraphe6.Font.Style;
  THEdit(getControl('TPARNIV6')).Font.size  :=  TheCourant.Paragraphe6.Font.size;
  THEdit(getControl('TPARNIV6')).Font.color  :=  TheCourant.Paragraphe6.Font.color;
  //
  THEdit(getControl('TPARNIV7')).Color :=  TheCourant.Paragraphe7.brush;
  THEdit(getControl('TPARNIV7')).Font.Name  :=  TheCourant.Paragraphe7.Font.Name;
  THEdit(getControl('TPARNIV7')).Font.Style  :=  TheCourant.Paragraphe7.Font.Style;
  THEdit(getControl('TPARNIV7')).Font.size  :=  TheCourant.Paragraphe7.Font.size;
  THEdit(getControl('TPARNIV7')).Font.color  :=  TheCourant.Paragraphe7.Font.color;
  //
  THEdit(getControl('TPARNIV8')).Color :=  TheCourant.Paragraphe8.brush;
  THEdit(getControl('TPARNIV8')).Font.Name  :=  TheCourant.Paragraphe8.Font.Name;
  THEdit(getControl('TPARNIV8')).Font.Style  :=  TheCourant.Paragraphe8.Font.Style;
  THEdit(getControl('TPARNIV8')).Font.size  :=  TheCourant.Paragraphe8.Font.size;
  THEdit(getControl('TPARNIV8')).Font.color  :=  TheCourant.Paragraphe8.Font.color;
  //
  THEdit(getControl('TPARNIV9')).Color :=  TheCourant.Paragraphe9.brush;
  THEdit(getControl('TPARNIV9')).Font.Name  :=  TheCourant.Paragraphe9.Font.Name;
  THEdit(getControl('TPARNIV9')).Font.Style  :=  TheCourant.Paragraphe9.Font.Style;
  THEdit(getControl('TPARNIV9')).Font.size  :=  TheCourant.Paragraphe9.Font.size;
  THEdit(getControl('TPARNIV9')).Font.color  :=  TheCourant.Paragraphe9.Font.color;
  //
  THEdit(getControl('TPAROUV')).Color :=  TheCourant.Ouvrage.brush;
  THEdit(getControl('TPAROUV')).Font.Name  :=  TheCourant.Ouvrage.Font.Name;
  THEdit(getControl('TPAROUV')).Font.Style  :=  TheCourant.Ouvrage.Font.Style;
  THEdit(getControl('TPAROUV')).Font.size  :=  TheCourant.Ouvrage.Font.size;
  THEdit(getControl('TPAROUV')).Font.color  :=  TheCourant.Ouvrage.Font.color;
  //
  THEdit(getControl('TPARVAR')).Color :=  TheCourant.Variante.brush;
  THEdit(getControl('TPARVAR')).Font.Name  :=  TheCourant.Variante.Font.Name;
  THEdit(getControl('TPARVAR')).Font.Style  :=  TheCourant.Variante.Font.Style;
  THEdit(getControl('TPARVAR')).Font.size  :=  TheCourant.Variante.Font.size;
  THEdit(getControl('TPARVAR')).Font.color  :=  TheCourant.Variante.Font.color;
  //
  THEdit(getControl('TPARDETOUV')).Color :=  TheCourant.SousDetail.brush;
  THEdit(getControl('TPARDETOUV')).Font.Name  :=  TheCourant.SousDetail.Font.Name;
  THEdit(getControl('TPARDETOUV')).Font.Style  :=  TheCourant.SousDetail.Font.Style;
  THEdit(getControl('TPARDETOUV')).Font.size  :=  TheCourant.SousDetail.Font.size;
  THEdit(getControl('TPARDETOUV')).Font.color  :=  TheCourant.SousDetail.Font.color;
  //
  THEdit(getControl('TPARLIG')).Color :=  TheCourant.Ligne.brush;
  THEdit(getControl('TPARLIG')).Font.Name  :=  TheCourant.Ligne.Font.Name;
  THEdit(getControl('TPARLIG')).Font.Style  :=  TheCourant.Ligne.Font.Style;
  THEdit(getControl('TPARLIG')).Font.size  :=  TheCourant.Ligne.Font.size;
  THEdit(getControl('TPARLIG')).Font.color  :=  TheCourant.Ligne.Font.color;
  //
  THEdit(getControl('TPARSTOT')).Color :=  TheCourant.SousTotaux.brush;
  THEdit(getControl('TPARSTOT')).Font.Name  :=  TheCourant.SousTotaux.Font.Name;
  THEdit(getControl('TPARSTOT')).Font.Style  :=  TheCourant.SousTotaux.Font.Style;
  THEdit(getControl('TPARSTOT')).Font.size  :=  TheCourant.SousTotaux.Font.size;
  THEdit(getControl('TPARSTOT')).Font.color  :=  TheCourant.SousTotaux.Font.color;
  //
  THEdit(getControl('TPARCOM')).Color :=  TheCourant.Commentaire.brush;
  THEdit(getControl('TPARCOM')).Font.Name  :=  TheCourant.Commentaire.Font.Name;
  THEdit(getControl('TPARCOM')).Font.Style  :=  TheCourant.Commentaire.Font.Style;
  THEdit(getControl('TPARCOM')).Font.size  :=  TheCourant.Commentaire.Font.size;
  THEdit(getControl('TPARCOM')).Font.color  :=  TheCourant.Commentaire.Font.color;
	//
  THSpinEdit(GetControl('BD2_DECALSDETAIL')).Value := TheCourant.DecalSousDetail;
  THSpinEdit(GetControl('BD2_DECALPARAG')).Value := TheCourant.DecalParag;
end;

procedure TOF_BTPARAFFDOC.SetEvents;
begin
	TToolbarButton97 (getControl('BBRUSHNIV1')).OnClick := Brushniv1Click;
	TToolbarButton97 (getControl('BFONTNIV1')).OnClick := Fontniv1Click;
	TToolbarButton97 (getControl('BCOLORNIV1')).OnClick := Colorniv1Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV2')).OnClick := Brushniv2Click;
	TToolbarButton97 (getControl('BFONTNIV2')).OnClick := Fontniv2Click;
	TToolbarButton97 (getControl('BCOLORNIV2')).OnClick := Colorniv2Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV3')).OnClick := Brushniv3Click;
	TToolbarButton97 (getControl('BFONTNIV3')).OnClick := Fontniv3Click;
	TToolbarButton97 (getControl('BCOLORNIV3')).OnClick := Colorniv3Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV4')).OnClick := Brushniv4Click;
	TToolbarButton97 (getControl('BFONTNIV4')).OnClick := Fontniv4Click;
	TToolbarButton97 (getControl('BCOLORNIV4')).OnClick := Colorniv4Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV5')).OnClick := Brushniv5Click;
	TToolbarButton97 (getControl('BFONTNIV5')).OnClick := Fontniv5Click;
	TToolbarButton97 (getControl('BCOLORNIV5')).OnClick := Colorniv5Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV6')).OnClick := Brushniv6Click;
	TToolbarButton97 (getControl('BFONTNIV6')).OnClick := Fontniv6Click;
	TToolbarButton97 (getControl('BCOLORNIV6')).OnClick := Colorniv6Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV7')).OnClick := Brushniv7Click;
	TToolbarButton97 (getControl('BFONTNIV7')).OnClick := Fontniv7Click;
	TToolbarButton97 (getControl('BCOLORNIV7')).OnClick := Colorniv7Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV8')).OnClick := Brushniv8Click;
	TToolbarButton97 (getControl('BFONTNIV8')).OnClick := Fontniv8Click;
	TToolbarButton97 (getControl('BCOLORNIV8')).OnClick := Colorniv8Click;
  //
	TToolbarButton97 (getControl('BBRUSHNIV9')).OnClick := Brushniv9Click;
	TToolbarButton97 (getControl('BFONTNIV9')).OnClick := Fontniv9Click;
	TToolbarButton97 (getControl('BCOLORNIV9')).OnClick := Colorniv9Click;
  //
	TToolbarButton97 (getControl('BBRUSHOUV')).OnClick := BrushOuvClick;
	TToolbarButton97 (getControl('BFONTOUV')).OnClick := FontOuvClick;
	TToolbarButton97 (getControl('BCOLOROUV')).OnClick := ColorOuvClick;
  //
	TToolbarButton97 (getControl('BBRUSHDETOUV')).OnClick := BrushDetOuvClick;
	TToolbarButton97 (getControl('BFONTDETOUV')).OnClick := FontDetOuvClick;
	TToolbarButton97 (getControl('BCOLORDETOUV')).OnClick := ColorDetOuvClick;
  //
	TToolbarButton97 (getControl('BBRUSHVAR')).OnClick := BrushVarClick;
	TToolbarButton97 (getControl('BFONTVAR')).OnClick := FontVarClick;
	TToolbarButton97 (getControl('BCOLORVAR')).OnClick := ColorVarClick;
  //
	TToolbarButton97 (getControl('BBRUSHLIG')).OnClick := BrushLigClick;
	TToolbarButton97 (getControl('BFONTLIG')).OnClick := FontLigClick;
	TToolbarButton97 (getControl('BCOLORLIG')).OnClick := ColorLigClick;
  //
	TToolbarButton97 (getControl('BBRUSHCOM')).OnClick := BrushComClick;
	TToolbarButton97 (getControl('BFONTCOM')).OnClick := FontComClick;
	TToolbarButton97 (getControl('BCOLORCOM')).OnClick := ColorComClick;
  //
	TToolbarButton97 (getControl('BBRUSHSTOT')).OnClick := BrushSTotClick;
	TToolbarButton97 (getControl('BFONTSTOT')).OnClick := FontSTotClick;
	TToolbarButton97 (getControl('BCOLORSTOT')).OnClick := ColorSTotClick;
  //
  THSpinEdit (GetControl('BD2_DECALSDETAIL')).OnChange := DESDETAILChange;
  THSpinEdit (GetControl('BD2_DECALPARAG')).OnChange := DEPARAGChange;

end;

procedure TOF_BTPARAFFDOC.Colorniv1Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV1')),Ecran) then
  begin
		TheCourant.Paragraphe1.Font.Color := THedit(getControl('TPARNIV1')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv1Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV1')),Ecran) then
  begin
		TheCourant.Paragraphe1.Font.Size   := THedit(getControl('TPARNIV1')).Font.Size;
		TheCourant.Paragraphe1.Font.Name   := THedit(getControl('TPARNIV1')).Font.Name;
		TheCourant.Paragraphe1.Font.Style   := THedit(getControl('TPARNIV1')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv1Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV1')),Ecran) then
  begin
		TheCourant.Paragraphe1.brush := THedit(getControl('TPARNIV1')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv2Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV2')),Ecran) then
  begin
		TheCourant.Paragraphe2.Font.Color := THedit(getControl('TPARNIV2')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv2Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV2')),Ecran) then
  begin
		TheCourant.Paragraphe2.Font.Size   := THedit(getControl('TPARNIV2')).Font.Size;
		TheCourant.Paragraphe2.Font.Name   := THedit(getControl('TPARNIV2')).Font.Name;
		TheCourant.Paragraphe2.Font.Style   := THedit(getControl('TPARNIV2')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv2Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV2')),Ecran) then
  begin
		TheCourant.Paragraphe2.brush := THedit(getControl('TPARNIV2')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv3Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV3')),Ecran) then
  begin
		TheCourant.Paragraphe3.Font.color := THedit(getControl('TPARNIV3')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv3Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV3')),Ecran) then
  begin
		TheCourant.Paragraphe3.Font.Size   := THedit(getControl('TPARNIV3')).Font.Size;
		TheCourant.Paragraphe3.Font.Name   := THedit(getControl('TPARNIV3')).Font.Name;
		TheCourant.Paragraphe3.Font.Style   := THedit(getControl('TPARNIV3')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv3Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV3')),Ecran) then
  begin
		TheCourant.Paragraphe3.brush := THedit(getControl('TPARNIV3')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv4Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV4')),Ecran) then
  begin
		TheCourant.Paragraphe4.Font.color := THedit(getControl('TPARNIV4')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv4Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV4')),Ecran) then
  begin
		TheCourant.Paragraphe4.Font.Size   := THedit(getControl('TPARNIV4')).Font.Size;
		TheCourant.Paragraphe4.Font.Name   := THedit(getControl('TPARNIV4')).Font.Name;
		TheCourant.Paragraphe4.Font.Style   := THedit(getControl('TPARNIV4')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv4Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV4')),Ecran) then
  begin
		TheCourant.Paragraphe5.brush := THedit(getControl('TPARNIV5')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv5Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV5')),Ecran) then
  begin
		TheCourant.Paragraphe5.Font.color := THedit(getControl('TPARNIV5')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv5Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV5')),Ecran) then
  begin
		TheCourant.Paragraphe5.Font.Size   := THedit(getControl('TPARNIV5')).Font.Size;
		TheCourant.Paragraphe5.Font.Name   := THedit(getControl('TPARNIV5')).Font.Name;
		TheCourant.Paragraphe5.Font.Style   := THedit(getControl('TPARNIV5')).Font.Style;

  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv5Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV5')),Ecran) then
  begin
		TheCourant.Paragraphe5.brush := THedit(getControl('TPARNIV5')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv6Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV6')),Ecran) then
  begin
		TheCourant.Paragraphe6.Font.color := THedit(getControl('TPARNIV6')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv6Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV6')),Ecran) then
  begin
		TheCourant.Paragraphe6.Font.Size   := THedit(getControl('TPARNIV6')).Font.Size;
		TheCourant.Paragraphe6.Font.Name   := THedit(getControl('TPARNIV6')).Font.Name;
		TheCourant.Paragraphe6.Font.Style   := THedit(getControl('TPARNIV6')).Font.Style;

  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv6Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV6')),Ecran) then
  begin
		TheCourant.Paragraphe6.brush := THedit(getControl('TPARNIV6')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv7Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV7')),Ecran) then
  begin
		TheCourant.Paragraphe7.Font.color := THedit(getControl('TPARNIV7')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv7Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV7')),Ecran) then
  begin
		TheCourant.Paragraphe7.Font.Size   := THedit(getControl('TPARNIV7')).Font.Size;
		TheCourant.Paragraphe8.Font.Name   := THedit(getControl('TPARNIV7')).Font.Name;
		TheCourant.Paragraphe7.Font.Style   := THedit(getControl('TPARNIV7')).Font.Style;

  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv7Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV7')),Ecran) then
  begin
		TheCourant.Paragraphe7.brush := THedit(getControl('TPARNIV7')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv8Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV8')),Ecran) then
  begin
		TheCourant.Paragraphe8.Font.color := THedit(getControl('TPARNIV8')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv8Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV8')),Ecran) then
  begin
		TheCourant.Paragraphe8.Font.Size   := THedit(getControl('TPARNIV8')).Font.Size;
		TheCourant.Paragraphe8.Font.Name   := THedit(getControl('TPARNIV8')).Font.Name;
		TheCourant.Paragraphe8.Font.Style   := THedit(getControl('TPARNIV8')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv8Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV8')),Ecran) then
  begin
		TheCourant.Paragraphe8.brush := THedit(getControl('TPARNIV8')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.Colorniv9Click (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARNIV9')),Ecran) then
  begin
		TheCourant.Paragraphe9.Font.color := THedit(getControl('TPARNIV9')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.Fontniv9Click (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARNIV9')),Ecran) then
  begin
		TheCourant.Paragraphe9.Font.Size   := THedit(getControl('TPARNIV9')).Font.Size;
		TheCourant.Paragraphe9.Font.Name   := THedit(getControl('TPARNIV9')).Font.Name;
		TheCourant.Paragraphe9.Font.Style   := THedit(getControl('TPARNIV9')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.Brushniv9Click(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARNIV9')),Ecran) then
  begin
		TheCourant.Paragraphe9.brush := THedit(getControl('TPARNIV9')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.ColorOuvClick (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPAROUV')),Ecran) then
  begin
		TheCourant.Ouvrage.Font.color := THedit(getControl('TPAROUV')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.FontOuvClick (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPAROUV')),Ecran) then
  begin
		TheCourant.Ouvrage.Font.Size   := THedit(getControl('TPAROUV')).Font.Size;
		TheCourant.Ouvrage.Font.Name   := THedit(getControl('TPAROUV')).Font.Name;
		TheCourant.Ouvrage.Font.Style   := THedit(getControl('TPAROUV')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.BrushOuvClick(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPAROUV')),Ecran) then
  begin
		TheCourant.Ouvrage.brush := THedit(getControl('TPAROUV')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.ColorVarClick (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARVAR')),Ecran) then
  begin
		TheCourant.Variante.Font.color := THedit(getControl('TPARVAR')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.ColorLigClick (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARLIG')),Ecran) then
  begin
		TheCourant.Ligne.Font.color := THedit(getControl('TPARLIG')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.ColorComClick (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARCOM')),Ecran) then
  begin
		TheCourant.Commentaire.Font.color := THedit(getControl('TPARCOM')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.ColorSTotClick (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARSTOT')),Ecran) then
  begin
		TheCourant.SousTotaux.Font.color := THedit(getControl('TPARSTOT')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.FontVarClick (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARVAR')),Ecran) then
  begin
		TheCourant.Variante.Font.Size   := THedit(getControl('TPARVAR')).Font.Size;
		TheCourant.Variante.Font.Name   := THedit(getControl('TPARVAR')).Font.Name;
		TheCourant.Variante.Font.Style   := THedit(getControl('TPARVAR')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.FontLigClick (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARLIG')),Ecran) then
  begin
		TheCourant.ligne.Font.Size   := THedit(getControl('TPARLIG')).Font.Size;
		TheCourant.Ligne.Font.Name   := THedit(getControl('TPARLIG')).Font.Name;
		TheCourant.Ligne.Font.Style   := THedit(getControl('TPARLIG')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.FontComClick (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARCOM')),Ecran) then
  begin
		TheCourant.Commentaire.Font.Size   := THedit(getControl('TPARCOM')).Font.Size;
		TheCourant.Commentaire.Font.Name   := THedit(getControl('TPARCOM')).Font.Name;
		TheCourant.Commentaire.Font.Style   := THedit(getControl('TPARCOM')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.FontStotClick (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARSTOT')),Ecran) then
  begin
		TheCourant.SousTotaux.Font.Size   := THedit(getControl('TPARSTOT')).Font.Size;
		TheCourant.SousTotaux.Font.Name   := THedit(getControl('TPARSTOT')).Font.Name;
		TheCourant.SousTotaux.Font.Style   := THedit(getControl('TPARSTOT')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.BrushVarClick(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARVAR')),Ecran) then
  begin
		TheCourant.Variante.brush := THedit(getControl('TPARVAR')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.ColorDetOuvClick (Sender : TObject);
begin
	if SelectFontColor(THedit(getControl('TPARDETOUV')),Ecran) then
  begin
		TheCourant.SousDetail.Font.color := THedit(getControl('TPARDETOUV')).Font.color;
  end;
end;

procedure TOF_BTPARAFFDOC.FontDetOuvClick (Sender : TObject);
begin
	if SelectFonte(THedit(getControl('TPARDETOUV')),Ecran) then
  begin
		TheCourant.SousDetail.Font.Size   := THedit(getControl('TPARDETOUV')).Font.Size;
		TheCourant.SousDetail.Font.Name   := THedit(getControl('TPARDETOUV')).Font.Name;
		TheCourant.SousDetail.Font.Style   := THedit(getControl('TPARDETOUV')).Font.Style;
  end;
end;

procedure TOF_BTPARAFFDOC.BrushDetOuvClick(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARDETOUV')),Ecran) then
  begin
		TheCourant.SousDetail.brush := THedit(getControl('TPARDETOUV')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.BrushLigClick(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARLIG')),Ecran) then
  begin
		TheCourant.Ligne.brush := THedit(getControl('TPARLIG')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.BrushComClick(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARCOM')),Ecran) then
  begin
		TheCourant.Commentaire.brush := THedit(getControl('TPARCOM')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.BrushStotClick(Sender: Tobject);
begin
	if SelectBrushColor (THedit(getControl('TPARSTOT')),Ecran) then
  begin
		TheCourant.SousTotaux.brush := THedit(getControl('TPARSTOT')).color;
  end;
end;

procedure TOF_BTPARAFFDOC.SetScreenToNormal;
begin
	TheCourant.Enregistrevaleurs;
end;

procedure TOF_BTPARAFFDOC.DESDETAILChange(Sender: Tobject);
begin
	TheCourant.DecalSousDetail := THSpinEdit (getControl('BD2_DECALSDETAIL')).Value;
end;

procedure TOF_BTPARAFFDOC.DEPARAGChange(Sender: Tobject);
begin
	TheCourant.DecalParag := THSpinEdit (getControl('BD2_DECALPARAG')).Value;
end;

Initialization
  registerclasses ( [ TOF_BTPARAFFDOC ] ) ;
end.
