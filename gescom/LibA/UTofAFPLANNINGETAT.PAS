{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPLANNINGETAT ()
Mots clefs ... : TOF;AFPLANNINGETAT
*****************************************************************}
Unit UTofAFPLANNINGETAT;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1,
     HMsgBox, UTOF, UTob, HTB97, AFDialog, extctrls, graphics,
     UtilPlanning, UtilTaches, DicoAF;

Type
  TOF_AFPLANNINGETAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private

      fListe              : TStringList;
      fTobParam           : TOB;  // tob des paramètres
      fStCodeParam        : String;

      fE_COULEURFOND      : THLabel;
      fE_FONTE            : THLabel;
      fE_ICONE            : TImage;
      fStIconeFileName    : String;

      fAction   : TActionFiche; // action sur la fiche

      procedure TobParamToScreen;
      procedure ScreenToTobParam;
      function pLoad (T:TOB): integer;
      function TestSaisie : Boolean;

      procedure vBtFondOnClick(SEnder: TObject);
      procedure vBtFonteOnClick(SEnder: TObject);
      procedure vBtIconeOnClick(SEnder: TObject);

      procedure vBtDeleteOnClick(SEnder: TObject);
      procedure vBtValiderOnClick(SEnder: TObject);
      procedure vBtInsertOnClick(SEnder: TObject);

  end;

const                       
	TexteMsgTache: array[1..6] of string 	= (
          {1}        'Suppression impossible. Cet état est utilisé dans un planning.',
          {2}        'Le code est obligatoire.',
          {3}        'Le Libellé est obligatoire.',
          {4}        'Ce code existe déjà.',
          {5}        'Erreur lors de la mise à jour.',
          {6}        'Suppression impossible. L''état Congés n''est pas supprimable.');


Implementation

procedure TOF_AFPLANNINGETAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGETAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGETAT.OnUpdate;
begin
  inherited;
end;

procedure TOF_AFPLANNINGETAT.OnLoad;
var
  vStream       : TStream;
  vST           : String;
  vQR           : TQuery;
  vTobParamXml  : TOB;

begin

  Inherited ;

  // recuperation des données
  vTobParamXml := TOB.Create('AFPARAMETAT',Nil,-1);

  vSt := 'select APE_ETATLIGNE, APE_LIBELLEETAT, APE_PARAMS from AFPLANNINGETAT where APE_ETATLIGNE = "' + fStCodeParam + '"';
  vQr := nil;
  try
    vQR := OpenSql(vSt,True);
    if Not Vqr.Eof then
      vTobParamXml.LoadDetailDB('AFPLANNINGETAT','', '', vQR, False, True) ;

    if vTobParamXml.Detail.Count>0 then
      begin
        // chargement de la liste
        fListe.Text := vTobParamXml.Detail[0].GetValue('APE_PARAMS');

        // transfert dans une stream
        vStream := TStringStream.Create(fListe.Text);

        // affichage de la clé et du libelle
        SetControlText('APE_ETATLIGNE', vTobParamXml.Detail[0].GetValue('APE_ETATLIGNE'));
        SetControlText('APE_LIBELLEETAT' , vTobParamXml.Detail[0].GetValue('APE_LIBELLEETAT'));

        // recuperation dans une tob virtuelle
        TOBLoadFromXMLStream(vStream,pLoad);

        // affichage a l'ecran des parametres
        TobParamToScreen;
        vStream.Free;

      end
    else
      begin
        SetControlText('APE_ETATLIGNE', '');
        SetControlText('APE_LIBELLEETAT', '');
      end;
  finally
    Ferme(vQR);
    vTobParamXml.Free;
  end;
end;

function TOF_AFPLANNINGETAT.pLoad (T:TOB): integer;
begin
  fTobParam.ClearDetail;
  fTobParam.Dupliquer(T,True,True);
  T.free;
  result := 0;
end;

procedure TOF_AFPLANNINGETAT.OnArgument (S : String ) ;
var
  Tmp       : String;
  champ     : String;
  valeur    : String;

begin

  inherited;

  // traitement des arguments
  fStCodeParam := '';
  Tmp:=(Trim(ReadTokenSt(S)));
  While (Tmp <>'') do
    Begin
      If Tmp<>'' then
        Begin
          DecodeArgument(Tmp, Champ, valeur);

          If Champ='CODE'   then fStCodeParam := valeur else
          If Champ='ACTION' then
            Begin
              If valeur='MODIFICATION' then
                begin
                  fAction := taModif;
                  SetControlEnabled('APE_ETATLIGNE', false);
                end
              else If valeur='CREATION' then
                begin
                  fAction := taCreat;
                  SetControlEnabled('APE_ETATLIGNE', true);
                end;
            End;
        End;
        Tmp:=(Trim(ReadTokenSt(S)));
    End;

  fListe := TStringList.Create;
  fTobParam := TOB.Create('les_etats',Nil,-1);

  TToolBarButton97(GetControl('BFond')).onClick := vBtFondOnClick;
  TToolBarButton97(GetControl('BFonte')).OnClick := vBtFonteOnClick;
  TToolBarButton97(GetControl('BIcone')).OnClick := vBtIconeOnClick;
  TToolBarButton97(GetControl('BDelete')).OnClick := vBtDeleteOnClick;
  TToolBarButton97(GetControl('BValider')).OnClick := vBtValiderOnClick;
  TToolBarButton97(GetControl('BInsert')).OnClick := vBtInsertOnClick;

  fE_COULEURFOND := THLabel(GetControl('E_COULEURFOND'));
  fE_FONTE := THLabel(GetControl('E_FONTE'));
  fE_ICONE := TImage(GetControl('E_ICONE'));
                  
  setControlVisible('E_ICONE', false);
  setControlVisible('BICONE', false);
  setControlVisible('TICONE', false);
end;                          

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGETAT.TobParamToScreen;
begin                            
  fE_COULEURFOND.Color := fTobParam.Detail[0].GetValue('E_COULEURFOND');
  fE_FONTE.Font.Color  := fTobParam.Detail[0].GetValue('E_COULEURFONTE');
  fE_FONTE.Font.Name   := fTobParam.Detail[0].GetValue('E_NOMFONTE');
  fE_FONTE.Font.Style  := EncodeFontStyle(fTobParam.Detail[0].GetValue('E_STYLEFONTE'));
  fE_FONTE.Font.Size   := fTobParam.Detail[0].GetValue('E_TAILLEFONTE');
  fE_ICONE.picture.LoadFromFile(fTobParam.Detail[0].GetValue('E_ICONE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2002
Modifié le ... :
Description .. : Recupere dans une TStrings les données saisies
Suite ........ : à l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGETAT.ScreenToTobParam;
var
  TobFille : TOB;

begin

  TobFille := TOB.Create('fille_param',fTobParam,-1);
  TobFille.AddChampSupValeur('E_COULEURFOND', fE_COULEURFOND.Color, false);
  TobFille.AddChampSupValeur('E_COULEURFONTE', fE_FONTE.Font.Color, false);
  TobFille.AddChampSupValeur('E_NOMFONTE', fE_FONTE.Font.Name, false);
  TobFille.AddChampSupValeur('E_STYLEFONTE', DecodeFontStyle(fE_FONTE.Font.Style), false);
  TobFille.AddChampSupValeur('E_TAILLEFONTE', fE_FONTE.Font.Size, false);
  TobFille.AddChampSupValeur('E_ICONE', fStIconeFileName, false);
end;

procedure TOF_AFPLANNINGETAT.vBtFondOnClick(SEnder: TObject);
begin
  fE_COULEURFOND.Color := execAFColorDialog;
end;

procedure TOF_AFPLANNINGETAT.vBtFonteOnClick(SEnder: TObject);
var
 vFont : TFont;

begin
  vFont := execAFFontDialog;
  fE_FONTE.Font.name := vFont.name;
  fE_FONTE.Font.size := vFont.size;
  fE_FONTE.Font.Style := vFont.style;
  fE_FONTE.Font.Color := vFont.Color;
end;

procedure TOF_AFPLANNINGETAT.vBtIconeOnClick(SEnder: TObject);
begin
  fStIconeFileName := execAFSavePictureDialog;
  fE_ICONE.Picture.LoadFromFile(fStIconeFileName);
end;       

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 27/08/2002
Modifié le ... :
Description .. : Verification de l'utilisation d'un etat avant la
                 suppression
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGETAT.vBtDeleteOnClick(SEnder: TObject);
var
  vSt : String;
  vQr : Tquery;

begin

  // on ne peux pas supprimer l'etat CONGES
  if GetControlText('APE_ETATLIGNE') = 'CON' then
    PGIBoxAF (TexteMsgTache[6],'')
  else                             
  begin           
    vSt := 'SELECT APL_NUMEROLIGNE FROM AFPLANNING WHERE APL_ETATLIGNE = "';
    vSt := vSt + GetControlText('APE_ETATLIGNE') + '"';

    if existeSql(vSt) then
      PGIBoxAF (TexteMsgTache[1],'')
    else
      begin
        vSt := 'DELETE FROM AFPLANNINGETAT WHERE APE_ETATLIGNE = "' + GetControlText('APE_ETATLIGNE') + '"';
        executeSql(vSt);

        //On se positionne sur le premier enregistrement
        vSt := 'SELECT APE_ETATLIGNE FROM AFPLANNINGETAT';
        vQR := nil;
        try
          vQR := openSql(vSt, true);
          if not vQR.eof then
            fStCodeParam := vQr.findField('APE_ETATLIGNE').AsString
          else
            fStCodeParam := '';
          self.Load;

        finally
          ferme(vQr);
        end;
      end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 27/08/2002
Modifié le ... :
Description .. : creation
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGETAT.vBtInsertOnClick(SEnder: TObject);
begin
  fAction := taCreat;
  SetControlEnabled('APE_ETATLIGNE', true);
  SetControlText('APE_ETATLIGNE','');
  SetControlText('APE_LIBELLEETAT','');
  New;
end;

// controle de l'existance du code
procedure TOF_AFPLANNINGETAT.vBtValiderOnClick(SEnder: TObject);
var
  vStream  : TStream;
  vTob     : Tob;

begin

  Inherited;

  if TestSaisie then
    begin
      vTob := TOB.Create('AFPLANNINGETAT',Nil,-1);
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
        vTob.PutValue('APE_ETATLIGNE', GetControlText('APE_ETATLIGNE'));
        vTob.PutValue('APE_LIBELLEETAT', GetControlText('APE_LIBELLEETAT'));
        vTob.PutValue('APE_PARAMS', fListe.Text);

        if (fAction = taCreat) then
          begin
            if vTob.InsertDB(nil, false) then
              ecran.close
            else
              PGIBoxAF (TexteMsgTache[4],'');
          end
        else
          begin
            if vTob.UpdateDB(false) then
              ecran.close
            else
              PGIBoxAF (TexteMsgTache[5],'');
          end;

      Finally
        vStream.Free;
        vTob.Free;
      end;
    end;
end;

function TOF_AFPLANNINGETAT.TestSaisie : Boolean;
begin
  result := true;
  if GetControlText('APE_ETATLIGNE') = '' then
    begin
      PGIBoxAF (TexteMsgTache[2],'');
      SetFocusControl ('APE_ETATLIGNE');
      result := false;
    end
  else if GetControlText('APE_LIBELLEETAT') = '' then
    begin
      PGIBoxAF (TexteMsgTache[3],'');
      SetFocusControl ('APE_LIBELLEETAT');
      result := false;
    end;
end;

procedure TOF_AFPLANNINGETAT.OnClose;
begin
  fTobParam.free;
  fListe.Free;
  Inherited;
end;

Initialization
  registerclasses ( [ TOF_AFPLANNINGETAT ] ) ;
end.
