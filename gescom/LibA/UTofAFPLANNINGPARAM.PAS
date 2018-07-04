{***********UNITE*************************************************
Auteur  ...... : CB
Créé le ...... : 03/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPARAMPLANNING ()
Mots clefs ... : TOF;AFPARAMPLANNING
*****************************************************************}
Unit UTofAFPLANNINGPARAM;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     windows,
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     uTob,
     HTB97,               
     AFDialog,
     graphics;

Type
  TOF_AFPLANNINGPARAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private

      fListe              : TStringList;
      fTobParam           : TOB;  // tob des paramètres
      fTobParamXml        : TOB;  // tob de la table (parametres en xml)
      fStCodeParam        : String;

      fLFOND        : THLabel;
      fLFONTE       : THLabel;
      fLSELECTION   : THLabel;
      fLSAMEDI      : THLabel;
      fLDIMANCHE    : THLabel;
      fLJOURSFERIES : THLabel;

      procedure TobParamToScreen;
      procedure ScreenToTobParam;
      function pLoad (T:TOB): integer;

      procedure vBFondOnClick(SEnder: TObject);
      procedure vBSelectionOnClick(SEnder: TObject);
      procedure vBCouleurSamediOnClick(SEnder: TObject);
      procedure vBCouleurDimancheOnClick(SEnder: TObject);
      procedure vBJoursFeriesOnClick(SEnder: TObject);
      procedure vBFonteOnClick(SEnder: TObject);

  end;

Implementation

procedure TOF_AFPLANNINGPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGPARAM.OnUpdate ;
var
  vStream  : TStream;

begin

  Inherited;

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

    // sauvegarde en base
    TOB.Create('fille_param',fTobParamXml,-1);

    fTobParamXml.Detail[0].PutValue('APP_CODEPARAM', GetControlText('APP_CODEPARAM'));
    fTobParamXml.Detail[0].PutValue('APP_LIBELLEPARAM', GetControlText('APP_LIBELLEPARAM'));
    fTobParamXml.Detail[0].PutValue('APP_PARAMS', fListe.Text) ;

    fTobParamXml.InsertOrUpdateDB ;

  Finally
    vStream.Free;
  end;

end;


procedure TOF_AFPLANNINGPARAM.OnLoad ;
var
  vStream    : TStream;
  vST        : String;
  vQR        : TQuery;

begin

  Inherited ;

  // recuperation des données
  fTobParamXml.ClearDetail ;
  vSt := 'select APP_CODEPARAM, APP_LIBELLEPARAM, APP_PARAMS from AFPLANNINGPARAM where APP_CODEPARAM = "' + fStCodeParam + '"';
  vQr := nil;
  try
    vQR := OpenSql(vSt,True);
    if Not Vqr.Eof then
      fTobParamXml.LoadDetailDB('AFPLANNINGPARAM','', '', vQR, False, True) ;

    if FTobParamXml.Detail.Count>0 then
      begin
        // chargement de la liste
        fListe.Text := fTobParamXml.Detail[0].GetValue('APP_PARAMS');

        // transfert dans une stream
        vStream := TStringStream.Create(fListe.Text);

        // affichage de la clé et du libelle
        SetControlText('APP_CODEPARAM', fTobParamXml.Detail[0].GetValue('APP_CODEPARAM'));
        SetControlText('APP_LIBELLEPARAM' , fTobParamXml.Detail[0].GetValue('APP_LIBELLEPARAM'));

        // recuperation dans une tob virtuelle
        TOBLoadFromXMLStream(vStream,pLoad);

        // affichage a l'ecran des parametres
        TobParamToScreen;
        vStream.Free;

      end;
  finally
    Ferme(vQR);
  end;
end;

function TOF_AFPLANNINGPARAM.pLoad (T:TOB): integer;
begin
  fTobParam.ClearDetail;
  fTobParam.Dupliquer(T,True,True);
  T.free;
  result := 0;
end;

procedure TOF_AFPLANNINGPARAM.OnArgument (S : String );
var
  vBFond            : TToolBarButton97;
  vBFonte           : TToolBarButton97;
  vBSelection       : TToolBarButton97;
  vBCouleurSamedi   : TToolBarButton97;
  vBCouleurDimanche : TToolBarButton97;
  vBJoursFeries     : TToolBarButton97;

begin

  Inherited;

  fStCodeParam := (Trim(ReadTokenSt(S)));
  fListe := TStringList.Create;
  fTobParam := TOB.Create('les_parametres',Nil,-1);
  fTobParamXml := TOB.Create('AFPLANNINGPARAM',Nil,-1);  //AB-030404

  vBFond := TToolBarButton97(GetControl('BFond'));
  vBFond.onClick := vBFondOnClick;

  vBFonte := TToolBarButton97(GetControl('BFonte'));
  vBFonte.OnClick := vBFonteOnClick;

  vBSelection := TToolBarButton97(GetControl('BSelection'));
  vBSelection.OnClick := vBSelectionOnClick;

  vBCouleurSamedi := TToolBarButton97(GetControl('BCouleurSamedi'));
  vBCouleurSamedi.OnClick := vBCouleurSamediOnClick;

  vBCouleurDimanche := TToolBarButton97(GetControl('BCouleurDimanche'));
  vBCouleurDimanche.OnClick := vBCouleurDimancheOnClick;

  vBJoursFeries := TToolBarButton97(GetControl('BCouleurJoursFeries'));
  vBJoursFeries.OnClick := vBJoursFeriesOnClick;

  fLFOND := THLabel(GetControl('LFOND'));
  fLFONTE := THLabel(GetControl('LFONTE'));
  fLSELECTION := THLabel(GetControl('LSELECTION'));
  fLSAMEDI := THLabel(GetControl('LSAMEDI'));
  fLDIMANCHE := THLabel(GetControl('LDIMANCHE'));
  fLJOURSFERIES := THLabel(GetControl('LJOURSFERIES'));

  setControlVisible('CADENCEMENT', false);
  setControlVisible('TCADENCEMENT', false);

end;

procedure TOF_AFPLANNINGPARAM.OnClose;
begin
  fTobParam.free;
  fTobParamXml.free;
  fListe.Free;
  Inherited;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGPARAM.TobParamToScreen;
begin
  SetControlText('FORMEGRAPHIQUE', fTobParam.Detail[0].GetValue('FORMEGRAPHIQUE'));
  SetControlText('CADENCEMENT', fTobParam.Detail[0].GetValue('CADENCEMENT'));
  SetControlText('FORMATDATE', fTobParam.Detail[0].GetValue('FORMATDATE'));

  if fTobParam.Detail[0].GetValue('LARGEURCOL') <> '' then
    SetControlText('LARGEURCOL', fTobParam.Detail[0].GetValue('LARGEURCOL'));

  if fTobParam.Detail[0].GetValue('LARGEURCOLENTETE') <> '' then
    SetControlText('LARGEURCOLENTETE', fTobParam.Detail[0].GetValue('LARGEURCOLENTETE'));

  fLSELECTION.Color := fTobParam.Detail[0].GetValue('LSELECTION');
  fLSAMEDI.Color := fTobParam.Detail[0].GetValue('LSAMEDI');
  fLDIMANCHE.Color := fTobParam.Detail[0].GetValue('LDIMANCHE');
  fLJOURSFERIES.Color := fTobParam.Detail[0].GetValue('LJOURSFERIES');

  fLFOND.Color := fTobParam.Detail[0].GetValue('LFOND');
  fLFONTE.Font.Name := fTobParam.Detail[0].GetValue('LFONTNAME');
  fLFONTE.Font.Color := fTobParam.Detail[0].GetValue('LFONTCOLOR');
  fLFONTE.Font.Size := fTobParam.Detail[0].GetValue('LFONTSIZE');

  fLFONTE.Font.Style := [];
  if pos('I', fTobParam.Detail[0].GetValue('LFONTSTYLE')) <> 0 then
  fLFONTE.Font.Style := [fsItalic];
  if pos('B', fTobParam.Detail[0].GetValue('LFONTSTYLE')) <> 0 then
  fLFONTE.Font.Style := fLFONTE.Font.Style + [fsBold];
  if pos('U', fTobParam.Detail[0].GetValue('LFONTSTYLE')) <> 0 then
  fLFONTE.Font.Style := fLFONTE.Font.Style + [fsUnderline];
  if pos('S', fTobParam.Detail[0].GetValue('LFONTSTYLE')) <> 0 then
  fLFONTE.Font.Style := fLFONTE.Font.Style + [fsStrikeout];
                                       
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGPARAM.ScreenToTobParam;
var
  TobFille : TOB;
  vStStyle : String;
begin

  TobFille := TOB.Create('fille_param',fTobParam,-1);
  TobFille.AddChampSupValeur('FORMEGRAPHIQUE', GetControlText('FORMEGRAPHIQUE'), false);
  TobFille.AddChampSupValeur('LARGEURCOL', GetControlText('LARGEURCOL'), false);
  TobFille.AddChampSupValeur('LARGEURCOLENTETE', GetControlText('LARGEURCOLENTETE'), false);
  TobFille.AddChampSupValeur('CADENCEMENT', GetControlText('CADENCEMENT'), false);
  TobFille.AddChampSupValeur('FORMATDATE', GetControlText('FORMATDATE'), false);
  TobFille.AddChampSupValeur('LFOND', fLFOND.Color, false);
  TobFille.AddChampSupValeur('LFONTNAME', fLFONTE.Font.Name, false);
  TobFille.AddChampSupValeur('LFONTCOLOR', fLFONTE.Font.Color, false);
  TobFille.AddChampSupValeur('LFONTSIZE', fLFONTE.Font.Size, false);
   
  vStStyle := '';
  if fsBold	in fLFONTE.Font.style then vStStyle := 'B';
  if fsItalic in fLFONTE.Font.style then vStStyle := vStStyle + 'I';
  if fsUnderline in fLFONTE.Font.style then vStStyle := vStStyle + 'U';
  if fsStrikeout in fLFONTE.Font.style then vStStyle := vStStyle + 'S';
  TobFille.AddChampSupValeur('LFONTSTYLE', vStStyle, false);

  TobFille.AddChampSupValeur('LSELECTION', fLSELECTION.Color, false);
  TobFille.AddChampSupValeur('LSAMEDI', fLSAMEDI.Color, false);
  TobFille.AddChampSupValeur('LDIMANCHE', fLDIMANCHE.Color, false);
  TobFille.AddChampSupValeur('LJOURSFERIES', fLJOURSFERIES.Color, false);

end;

procedure TOF_AFPLANNINGPARAM.vBFondOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execAFColorDialog;
  if vColor <> 0 then THLabel(GetControl('LFOND')).Color := vColor;
end;

procedure TOF_AFPLANNINGPARAM.vBSelectionOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execAFColorDialog;
  if vColor <> 0 then THLabel(GetControl('LSELECTION')).Color := vColor;
end;

procedure TOF_AFPLANNINGPARAM.vBCouleurSamediOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execAFColorDialog;
  if vColor <> 0 then THLabel(GetControl('LSAMEDI')).Color := vColor;
end;

procedure TOF_AFPLANNINGPARAM.vBCouleurDimancheOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execAFColorDialog;
  if vColor <> 0 then THLabel(GetControl('LDIMANCHE')).Color := vColor;
end;

procedure TOF_AFPLANNINGPARAM.vBJoursFeriesOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execAFColorDialog;
  THLabel(GetControl('LJOURSFERIES')).Color := vColor;
end;

procedure TOF_AFPLANNINGPARAM.vBFonteOnClick(SEnder: TObject);
var
 vFont : TFont;
                    
begin
  vFont := execAFFontDialog;
  fLFONTE.Font.name := vFont.name;
  fLFONTE.Font.size := vFont.size;
  fLFONTE.Font.Style := vFont.style;
  fLFONTE.Font.Color := vFont.Color;
  vFont.Free; 
end;

Initialization
  registerclasses ( [ TOF_AFPLANNINGPARAM ] ) ;

end.
