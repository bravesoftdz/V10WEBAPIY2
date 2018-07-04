Unit UTofRTParamPlanning;

Interface

Uses
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     sysutils,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     uTob,
     HTB97,
     RTDialog,
     graphics,
     Vierge,
     EntRT
     ;

Type
  TOF_RTParamPlanning = Class (TOF)
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
      fTobParamSauv       : TOB;  // Sauvegarde tob des paramètres
      fStCodeParam        : String;
      fStCodeIntervenant  : String;
      EnregModif          : Boolean;

      fLFOND        : THLabel;
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
      procedure EnregModifs;
      Function  CompareTobs(Tob1,tob2 : tob) : boolean;

  end;


Implementation

procedure TOF_RTParamPlanning.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTParamPlanning.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTParamPlanning.OnUpdate ;
begin
  Inherited;
  EnregModifs;
  EnregModif := True;
end;


procedure TOF_RTParamPlanning.OnLoad ;
var
  vStream    : TStream;
  vST        : String;
  vQR        : TQuery;
  TobTemp    : TOB;

begin

  Inherited ;

  // affichage de l'intervenant et de son nom
//  SetControlText('CODEINTERVENANT',fStCodeIntervenant);
  SetControlText('NOMINTERVENANT',RechDom('TTUtilisateur',fStCodeIntervenant,False)) ;
  SetControlEnabled('NOMINTERVENANT',False) ;

  // recuperation des données
  fTobParamXml.ClearDetail ;
  vSt := 'select RPP_CODEPARAMPLAN, RPP_INTERVENANT, RPP_LIBELLEPLAN, RPP_PARAMS from RTPARAMPLANNING where RPP_CODEPARAMPLAN = "' + fStCodeParam + '" and RPP_INTERVENANT = "' + fStCodeIntervenant + '"';
  vQr := nil;
  try
    vQR := OpenSql(vSt,True);
    if Not Vqr.Eof then
      fTobParamXml.LoadDetailDB('RTPARAMPLANNING','', '', vQR, False, True);

    if FTobParamXml.Detail.Count>0 then
      begin
        // chargement de la liste
        fListe.Text := fTobParamXml.Detail[0].GetValue('RPP_PARAMS');

        // transfert dans une stream
        vStream := TStringStream.Create(fListe.Text);


        // recuperation dans une tob virtuelle
        TOBLoadFromXMLStream(vStream,pLoad);

        // affichage a l'ecran des parametres
        TobParamToScreen;
        vStream.Free;

      end
    else
      begin
        TobTemp := Tob.create('fille_param',fTobParam,-1);
        TobTemp.AddChampSupValeur('FORMEGRAPHIQUE', 'PGA', false);
        TobTemp.AddChampSupValeur('LARGEURCOL', '80', false);
        TobTemp.AddChampSupValeur('HAUTLIGNEDATA', '20', false);
        TobTemp.AddChampSupValeur('CADENCEMENT', '006', false);
        TobTemp.AddChampSupValeur('HEUREDEBUT', '08:00', false);
        TobTemp.AddChampSupValeur('HEUREFIN', '18:00', false);
        TobTemp.AddChampSupValeur('TYPEACTIONS', TraduireMemoire('<<Tous>>'), false);
        TobTemp.AddChampSupValeur('LFOND', StringToColor('clWhite'), false);
        TobTemp.AddChampSupValeur('LSELECTION', StringToColor('16777088'), false);
        TobTemp.AddChampSupValeur('LSAMEDI', StringToColor('33023'), false);
        TobTemp.AddChampSupValeur('LDIMANCHE', StringToColor('33023'), false);
        TobTemp.AddChampSupValeur('LJOURSFERIES', StringToColor('33023'), false);
        TobParamToScreen;
      end;
  finally
    Ferme(vQR);
  end;
  fTobParamSauv.Dupliquer (fTobParam,True,True,False);

end;

function TOF_RTParamPlanning.pLoad (T:TOB): integer;
begin
  fTobParam.ClearDetail;
  fTobParam.Dupliquer(T,True,True);
  T.free;
  result := 0;
end;

procedure TOF_RTParamPlanning.OnArgument (S : String );
var
  vBFond            : TToolBarButton97;
  vBSelection       : TToolBarButton97;
  vBCouleurSamedi   : TToolBarButton97;
  vBCouleurDimanche : TToolBarButton97;
  vBJoursFeries     : TToolBarButton97;

begin

  Inherited;

  fStCodeParam := (Trim(ReadTokenSt(S)));
  fStCodeIntervenant := (Trim(ReadTokenSt(S)));
  fListe := TStringList.Create;
  fTobParam := TOB.Create('les_parametres',Nil,-1);
  fTobParamXml := TOB.Create('RTPARAMPLANNING',Nil,-1);
  fTobParamSauv := TOB.Create('les_parametres_Sauv',Nil,-1);

  vBFond := TToolBarButton97(GetControl('BFond'));
  vBFond.onClick := vBFondOnClick;

  vBSelection := TToolBarButton97(GetControl('BSelection'));
  vBSelection.OnClick := vBSelectionOnClick;

  vBCouleurSamedi := TToolBarButton97(GetControl('BCouleurSamedi'));
  vBCouleurSamedi.OnClick := vBCouleurSamediOnClick;

  vBCouleurDimanche := TToolBarButton97(GetControl('BCouleurDimanche'));
  vBCouleurDimanche.OnClick := vBCouleurDimancheOnClick;

  vBJoursFeries := TToolBarButton97(GetControl('BCouleurJoursFeries'));
  vBJoursFeries.OnClick := vBJoursFeriesOnClick;

  fLFOND := THLabel(GetControl('LFOND'));
  fLSELECTION := THLabel(GetControl('LSELECTION'));
  fLSAMEDI := THLabel(GetControl('LSAMEDI'));
  fLDIMANCHE := THLabel(GetControl('LDIMANCHE'));
  fLJOURSFERIES := THLabel(GetControl('LJOURSFERIES'));
  SetControlProperty ('LFOND','TRANSPARENT',False);
  SetControlProperty ('LSELECTION','TRANSPARENT',False);
  SetControlProperty ('LSAMEDI','TRANSPARENT',False);
  SetControlProperty ('LDIMANCHE','TRANSPARENT',False);
  SetControlProperty ('LJOURSFERIES','TRANSPARENT',False);

  EnregModif := False;
  if fStCodeParam = 'PL1' then
     begin
     SetControlVisible ('THPP_TAILLECOLONNE',False);
     SetControlVisible ('LARGEURCOL',False);
     SetControlVisible ('THPP_HAUTLIGNE',False);
     SetControlVisible ('HAUTLIGNEDATA',False);
     SetControlVisible ('THPP_FORMEGRAPHIQUE',False);
     SetControlVisible ('FORMEGRAPHIQUE',False);
     SetControlVisible ('THPP_COULEURFOND',False);
     SetControlVisible ('BFOND',False);
     SetControlVisible ('LFOND',False);
     SetControlVisible ('TCADENCEMENT',False);
     SetControlVisible ('CADENCEMENT',False);
     end;
  if fStCodeParam = 'PL2' then
     begin
     SetControlVisible ('THPP_HEUREDEBUT',False);
     SetControlVisible ('HEUREDEBUT',False);
     SetControlVisible ('THPP_HEUREFIN',False);
     SetControlVisible ('HEUREFIN',False);
     SetControlVisible ('TCADENCEMENT',False);
     SetControlVisible ('CADENCEMENT',False);
     end;
end;

procedure TOF_RTParamPlanning.OnClose;
begin
  if (EnregModif = False) then
     begin
     fTobParam.ClearDetail;
     ScreenToTobParam;
     if CompareTobs (fTobParam,fTobParamSauv) = False then
        Case PGIAskCancel(TraduireMemoire('Voulez-vous enregistrer les modifications ?'),TFVierge(ecran).Caption) of
          mrCancel :
              begin
              LastError:=1;
              exit;
              end;
          mrYes :
              begin
              EnregModifs;
              ChargeParamPlanning;
              end;
          mrNo :
        end ;
     end
  else ChargeParamPlanning;
  fTobParam.free;
  fTobParamXml.free;
  fListe.Free;
  fTobParamSauv.free;
  Inherited;
end ;

procedure TOF_RTParamPlanning.EnregModifs;
var vStream  : TStream;
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
    // sauvegarde en base
    if FTobParamXml.Detail.Count=0 then TOB.Create('RTPARAMPLANNING',fTobParamXml,-1);
    fTobParamXml.Detail[0].PutValue('RPP_CODEPARAMPLAN', fStCodeParam);
    fTobParamXml.Detail[0].PutValue('RPP_INTERVENANT', fStCodeIntervenant);
    fTobParamXml.Detail[0].PutValue('RPP_PARAMS', fListe.Text) ;
    fTobParamXml.InsertOrUpdateDB ;

  Finally
    vStream.Free;
  end;
end;

Function TOF_RTParamPlanning.CompareTobs(Tob1,tob2 : tob) : boolean;
var i : integer;
begin
Result:=True;
for i := 1000 to ((Tob1.Detail[0].ChampsSup.count-1)+1000) do
    begin
    if Tob1.Detail[0].GetValue(Tob1.Detail[0].GetNomChamp(i)) <> Tob2.Detail[0].GetValue(Tob2.Detail[0].GetNomChamp(i)) then
       begin
       Result:=False;
       exit;
       end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_RTParamPlanning.TobParamToScreen;
begin
  SetControlText('FORMEGRAPHIQUE', fTobParam.Detail[0].GetValue('FORMEGRAPHIQUE'));
  SetControlText('CADENCEMENT', fTobParam.Detail[0].GetValue('CADENCEMENT'));

  if fTobParam.Detail[0].GetValue('LARGEURCOL') <> '' then
    SetControlText('LARGEURCOL', fTobParam.Detail[0].GetValue('LARGEURCOL'));

  if fTobParam.Detail[0].GetValue('HAUTLIGNEDATA') <> '' then
    SetControlText('HAUTLIGNEDATA', fTobParam.Detail[0].GetValue('HAUTLIGNEDATA'));

  if fTobParam.Detail[0].GetValue('HEUREDEBUT') <> '' then
    SetControlText('HEUREDEBUT', fTobParam.Detail[0].GetValue('HEUREDEBUT'));

  if fTobParam.Detail[0].GetValue('HEUREFIN') <> '' then
    SetControlText('HEUREFIN', fTobParam.Detail[0].GetValue('HEUREFIN'));

  if fTobParam.Detail[0].GetValue('TYPEACTIONS') <> '' then
    SetControlText('TYPEACTIONS', fTobParam.Detail[0].GetValue('TYPEACTIONS'));

  fLSELECTION.Color := fTobParam.Detail[0].GetValue('LSELECTION');
  fLSAMEDI.Color := fTobParam.Detail[0].GetValue('LSAMEDI');
  fLDIMANCHE.Color := fTobParam.Detail[0].GetValue('LDIMANCHE');
  fLJOURSFERIES.Color := fTobParam.Detail[0].GetValue('LJOURSFERIES');
  fLFOND.Color := fTobParam.Detail[0].GetValue('LFOND');

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_RTParamPlanning.ScreenToTobParam;
var
  TobFille : TOB;
begin

  TobFille := TOB.Create('fille_param',fTobParam,-1);
  TobFille.AddChampSupValeur('FORMEGRAPHIQUE', GetControlText('FORMEGRAPHIQUE'), false);
  TobFille.AddChampSupValeur('LARGEURCOL', GetControlText('LARGEURCOL'), false);
  TobFille.AddChampSupValeur('HAUTLIGNEDATA', GetControlText('HAUTLIGNEDATA'), false);
  TobFille.AddChampSupValeur('CADENCEMENT', GetControlText('CADENCEMENT'), false);
  TobFille.AddChampSupValeur('HEUREDEBUT', GetControlText('HEUREDEBUT'), false);
  TobFille.AddChampSupValeur('HEUREFIN', GetControlText('HEUREFIN'), false);
  TobFille.AddChampSupValeur('TYPEACTIONS', GetControlText('TYPEACTIONS'), false);
  TobFille.AddChampSupValeur('LFOND', fLFOND.Color, false);
  TobFille.AddChampSupValeur('LSELECTION', fLSELECTION.Color, false);
  TobFille.AddChampSupValeur('LSAMEDI', fLSAMEDI.Color, false);
  TobFille.AddChampSupValeur('LDIMANCHE', fLDIMANCHE.Color, false);
  TobFille.AddChampSupValeur('LJOURSFERIES', fLJOURSFERIES.Color, false);

end;

procedure TOF_RTParamPlanning.vBFondOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execRTColorDialog(THLabel(GetControl('LFOND')).Color);
  if vColor <> 0 then THLabel(GetControl('LFOND')).Color := vColor;
end;

procedure TOF_RTParamPlanning.vBSelectionOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execRTColorDialog(THLabel(GetControl('LSELECTION')).Color);
  if vColor <> 0 then THLabel(GetControl('LSELECTION')).Color := vColor;
end;

procedure TOF_RTParamPlanning.vBCouleurSamediOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execRTColorDialog(THLabel(GetControl('LSAMEDI')).Color);
  if vColor <> 0 then THLabel(GetControl('LSAMEDI')).Color := vColor;
end;

procedure TOF_RTParamPlanning.vBCouleurDimancheOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execRTColorDialog(THLabel(GetControl('LDIMANCHE')).Color);
  if vColor <> 0 then THLabel(GetControl('LDIMANCHE')).Color := vColor;
end;

procedure TOF_RTParamPlanning.vBJoursFeriesOnClick(SEnder: TObject);
var
  vColor : TColor;
begin
  vColor := execRTColorDialog(THLabel(GetControl('LJOURSFERIES')).Color);
  if vColor <> 0 then THLabel(GetControl('LJOURSFERIES')).Color := vColor;
end;


Initialization
  registerclasses ( [ TOF_RTParamPlanning ] ) ;

end.
