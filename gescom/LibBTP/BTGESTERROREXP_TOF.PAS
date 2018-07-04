{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/12/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTGESTERROREXP ()
Mots clefs ... : TOF;BTGESTERROREXP
*****************************************************************}
Unit BTGESTERROREXP_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Graphics,
     Classes,
     ExtCtrls,
     Types,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AglInit,
     UGestionErreurExportLSE,
     UTOF ;

Type
  TOF_BTGESTERROREXP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    ImTypeArticle : TImageList;
  	LesErreurs : TlistErrorExport;
    IMG1,IMG2  : TImage;
    LISTERROR  : TlistBox;
    procedure ListDrawItem (Control: TWinControl; Index: Integer;Rect: TRect; State: TOwnerDrawState);
    procedure ChargeLaListe;
    procedure ListDblClick (Sender : Tobject);
  end ;

procedure TraitelesErreurs ( lesErreurs : TlistErrorExport );

Implementation

procedure TraitelesErreurs ( lesErreurs : TlistErrorExport );
var TOBP : TOB;
begin
  TOBP := TOB.Create ('LES PARAMS',nil,-1);
  TOBP.data := lesErreurs;
  TheTOB := TOBP;
  TRY
		AGLLanceFiche('Z','BTGESTERROR','','','');
  FINALLY
  	TOBP.free;
	END;
  TheTOB := nil;
end;


procedure TOF_BTGESTERROREXP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.OnArgument (S : String ) ;
var UneImage : Timage;
begin
  Inherited ;
  if LaTOB <> nil then LesErreurs := TlistErrorExport (LaTOB.data);
  IMG1 := Timage(GetControl('IMG1'));
  IMG2 := Timage(GetControl('IMG2'));
  LISTERROR := TListBox(getControl('LISTERROR'));
  LISTERROR.OnDrawItem := ListDrawItem;
  LISTERROR.Style := lbOwnerDrawFixed;
  LISTERROR.ItemHeight := 18;
  ImTypeArticle := TImageList.create (Tform(ecran));
  ImTypeArticle.height := 16;
  ImTypeArticle.Width := 16;
	UneImage := TImage(getControl('IMG1'));
  UneIMage.Transparent := true;
  ImTypeArticle.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
	UneImage := TImage(getControl('IMG2'));
  UneIMage.Transparent := true;
  ImTypeArticle.Add ( TBitMap(UneImage.Picture.Bitmap),nil);
  ChargeLaListe;
  LISTERROR.OnDblClick := ListDblClick;
end ;

procedure TOF_BTGESTERROREXP.OnClose ;
begin
  ImTypeArticle.free;
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTGESTERROREXP.ListDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  R: TRect;
  i: Integer;
  Canvas : Tcanvas;
begin
  with LISTERROR.Canvas do
  begin
    FillRect(Rect);
    Canvas := LISTERROR.Canvas;
    if LesErreurs.Items[Index].corrected then
    begin
    	ImTypeArticle.Draw(CanVas, Rect.left, Rect.top, 1)
    end else
    begin
    	ImTypeArticle.Draw(CanVas, Rect.left, Rect.top, 0);
    end;
    TextOut(Rect.Left+30,Rect.Top+2,LISTERROR.Items.ValueFromIndex[Index] );
  end;
end;

procedure TOF_BTGESTERROREXP.ChargeLaListe;
var Indice : integer;
begin
	for Indice := 0 to LesErreurs.Count -1 do
  begin
    if lesErreurs.Items[Indice].TypeErreur = TteArticle then
    begin
    	LISTERROR.AddItem(' Référence L.S.E non définie sur l''article '+lesErreurs.Items[Indice].CodeArticle,nil );
    end else
    begin
    	LISTERROR.AddItem(' Référence L.S.E non définie sur le client '+lesErreurs.Items[Indice].Tiers,nil );
    end;
  end;
end;

procedure TOF_BTGESTERROREXP.ListDblClick(Sender: Tobject);
begin
	if lesErreurs.items[LISTERROR.ItemIndex].TypeErreur = ttetiers then
  begin
    AGLLanceFiche('GC','GCTIERS','',lesErreurs.items[LISTERROR.ItemIndex].auxiliaire,
    							'T_NATUREAUXI='+lesErreurs.items[LISTERROR.ItemIndex].NatureAuxi+';ACTION=MODIFICATION');
    if VerifExportAuxiliaire (lesErreurs.items[LISTERROR.ItemIndex].auxiliaire) then
    begin
      lesErreurs.items[LISTERROR.ItemIndex].corrected := true;
    end else
    begin
      lesErreurs.items[LISTERROR.ItemIndex].corrected := false;
    end;
  end else
  begin
  	V_PGI.DispatchTT(7, taModif, lesErreurs.items[LISTERROR.ItemIndex].article, '', 'TYPEARTICLE=' + lesErreurs.items[LISTERROR.ItemIndex].Typearticle);
    if VerifExportArticle (lesErreurs.items[LISTERROR.ItemIndex].article) then
    begin
      lesErreurs.items[LISTERROR.ItemIndex].corrected := true;
    end else
    begin
      lesErreurs.items[LISTERROR.ItemIndex].corrected := false;
    end;
  end;
  LISTERROR.Refresh;
end;

Initialization
  registerclasses ( [ TOF_BTGESTERROREXP ] ) ; 
end.
