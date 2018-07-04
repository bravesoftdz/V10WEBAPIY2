{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : RTETATPLANNING (RTETATPLANNING)
Mots clefs ... : TOM;RTETATPLANNING
*****************************************************************}
Unit UtomRTEtatPlan ;

Interface

Uses Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
{$ENDIF}
     sysutils,
     HCtrls,
     UTOM,
     UTob,
     HTB97,
     RTDialog,
     graphics,
     LibPlanning,
     EntRT
      ;

Type
  TOM_RTETATPLANNING = Class (TOM)
    private

      fListe              : TStringList;
      fTobParam           : TOB;  // tob des paramètres
      fStCodeIntervenant  : String;
      fE_COULEURFOND      : THLabel;
      fE_FONTE            : THLabel;
      Modif               : Boolean;
      function pLoad (T:TOB): integer;
      procedure TobParamToScreen;
      procedure ScreenToTobParam;
      procedure vBtFondOnClick(SEnder: TObject);
      procedure vBtFonteOnClick(SEnder: TObject);
      procedure EnregModifs;

    Public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

uses rtfCounter;

procedure TOM_RTETATPLANNING.OnNewRecord ;
begin
  Inherited ;
  fE_COULEURFOND.Color := ClWhite;
  fE_FONTE.Font.name := 'Ms Sans Serif';
  fE_FONTE.Font.size := 8;
  fE_FONTE.Font.Style := [];
  fE_FONTE.Font.Color := ClBlack;
end ;

procedure TOM_RTETATPLANNING.OnDeleteRecord ;
begin
  Inherited ;
  Modif := True;
end ;

procedure TOM_RTETATPLANNING.OnUpdateRecord ;
begin
  Inherited ;

  // sauvegarde en base
  if (DS.State = dsInsert) and (GetField ('REP_CODEETATACT') <> '') then
    begin
    SetField('REP_INTERVENANT', fStCodeIntervenant);
    SetField('REP_LIBELLE',RechDom('RTETATACTION',GetField('REP_CODEETATACT'),FALSE));
    SetControlEnabled ('REP_CODEETATACT',False);
    EnregModifs;
    end;
  Modif := True;
end ;

procedure TOM_RTETATPLANNING.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_RTETATPLANNING.OnLoadRecord ;
var vStream       : TStream;
    st            : string;
begin
  Inherited ;
  if (DS<>nil) and ( not (ds.state in [dsinsert])) then
  begin
  // chargement de la liste
//EPZ  fListe.Text := GetField('REP_PARAMS');
  St := GetRtfStringText(GetField('REP_PARAMS'));

  // transfert dans une stream
//EPZ  vStream := TStringStream.Create(fListe.Text);
  vStream := TStringStream.Create(St);

  // recuperation dans une tob virtuelle
  TOBLoadFromXMLStream(vStream,pLoad);

  // affichage a l'ecran des parametres
  TobParamToScreen;
  vStream.Free;
  end;
end ;

procedure TOM_RTETATPLANNING.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_RTETATPLANNING.OnArgument ( S: String ) ;
begin
  Inherited ;
  fStCodeIntervenant := (Trim(ReadTokenSt(S)));
  fListe := TStringList.Create;
  fTobParam := TOB.Create('les_etats',Nil,-1);
  TToolBarButton97(GetControl('BFond')).onClick := vBtFondOnClick;
  TToolBarButton97(GetControl('BFonte')).OnClick := vBtFonteOnClick;
  fE_COULEURFOND := THLabel(GetControl('REP_COULEURFOND'));
  fE_FONTE := THLabel(GetControl('REP_FONTE'));
  SetControlProperty ('REP_COULEURFOND','TRANSPARENT',False);
  SetControlProperty ('REP_FONTE','TRANSPARENT',False);
  Modif := False;
end ;

procedure TOM_RTETATPLANNING.OnClose ;
begin
  Inherited ;
  fListe.Free;
  fTobParam.free;
  if Modif = True then ChargeParamEtatPlan;
end ;

procedure TOM_RTETATPLANNING.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_RTETATPLANNING.vBtFondOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execRTColorDialog (fE_COULEURFOND.Color);
  if vColor <> 0 then
     begin
      fE_COULEURFOND.Color := vColor;
      EnregModifs;
     end;
end;

procedure TOM_RTETATPLANNING.vBtFonteOnClick(SEnder: TObject);
var
 vFont : TFont;

begin
  vFont := execRTFontDialog (fE_FONTE.Font);
  fE_FONTE.Font.name := vFont.name;
  fE_FONTE.Font.size := vFont.size;
  fE_FONTE.Font.Style := vFont.style;
  fE_FONTE.Font.Color := vFont.Color;
  vFont.Free;
  EnregModifs;
end;

{***********A.G.L.***********************************************
Modifié le ... : copie la tob retournée par TOBLoadFromXMLStream
Description .. : dans la tob de travail
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TOM_RTETATPLANNING.pLoad (T:TOB): integer;
begin
  fTobParam.ClearDetail;
  fTobParam.Dupliquer(T,True,True);
  T.free;
  result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOM_RTETATPLANNING.TobParamToScreen;
begin                            
  fE_COULEURFOND.Color := fTobParam.Detail[0].GetValue('REP_COULEURFOND');
  fE_FONTE.Font.Color  := fTobParam.Detail[0].GetValue('REP_COULEURFONTE');
  fE_FONTE.Font.Name   := fTobParam.Detail[0].GetValue('REP_NOMFONTE');
  fE_FONTE.Font.Style  := EncodeFontStyle(fTobParam.Detail[0].GetValue('REP_STYLEFONTE'));
  fE_FONTE.Font.Size   := fTobParam.Detail[0].GetValue('REP_TAILLEFONTE');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOM_RTETATPLANNING.ScreenToTobParam;
var
  TobFille : TOB;

begin

  TobFille := TOB.Create('fille_param',fTobParam,-1);
  TobFille.AddChampSupValeur('REP_COULEURFOND', fE_COULEURFOND.Color, false);
  TobFille.AddChampSupValeur('REP_COULEURFONTE', fE_FONTE.Font.Color, false);
  TobFille.AddChampSupValeur('REP_NOMFONTE', fE_FONTE.Font.Name, false);
  TobFille.AddChampSupValeur('REP_STYLEFONTE', DecodeFontStyle(fE_FONTE.Font.Style), false);
  TobFille.AddChampSupValeur('REP_TAILLEFONTE', fE_FONTE.Font.Size, false);
end;

procedure TOM_RTETATPLANNING.EnregModifs;
var vStream  : TStream;
    st       : string;
begin
  Inherited ;
  vStream := TStringStream.Create('') ;
  Try
    fTobParam.ClearDetail;

    // recuperation dans une tob des données saisies : fTobParam
    ScreenToTobParam;

    // chargement dans tstream
    fTobParam.SaveToXmlStream(vStream,True,True);
    vStream.Seek(0,0) ;

    // transfert dans une liste
    fListe.LoadFromStream(vStream) ;

    if not(DS.State in [dsInsert,dsEdit]) then
      begin
      DS.edit; // pour passer DS.state en mode dsEdit
      end;
    st := GetRtfStringText(fListe.Text);
    SetField('REP_PARAMS', St);

  Finally
    vStream.Free;
  end;
end;

Initialization
  registerclasses ( [ TOM_RTETATPLANNING ] ) ; 
end.

