{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCONTROLECPTLIASSE ()
Mots clefs ... : TOF;CPCONTROLECPTLIASSE
*****************************************************************}
Unit CPTOTALCPTLIASSE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl,
{$ELSE}
     Fe_Main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$ENDIF}
     forms,
     HPanel,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,    // THSystemMenu
     ParamSoc,    // GetParamSocSecur
     ZCumul,      // TZCumul
     uLibWindows, // IIF
     uTobDebug,   // TobDebug
     CalcOle,     // GetCumul
     uTob,
     UTOF ;

Type
  TOF_CPTOTALCPTLIASSE = Class (TOF)

    FListe : THGrid;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnDblClickFListe (Sender : TObject);

  private

    FTobRub : Tob;
    FTobFListe : Tob;
    FZCumul : TZCumul;

    FStArgument      : string;
    FStGeneral       : string;
    FStLiasseDossier : string;
    FStDateDeb       : string;
    FStDateFin       : string;

    procedure TraiteInfoCompte(vTobRub : Tob; vSt: string);

  end;

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_CPTotalCptLiasse( vStArgument : string );
////////////////////////////////////////////////////////////////////////////////

Implementation

uses  CPSTATUTDOSSIER_TOF, // SupprimeDoublonAgregat(FTobRub);
      RUBRIQUE_TOM; // ParametrageRubrique

const cColAgregat = 1;
      cColLibelle = 2;
      cColTotalFiscal = 3;
      cColEcart = 4;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_CPTotalCptLiasse( vStArgument : string );
begin
  AGLLanceFiche('CP', 'CPTOTALCPTLIASSE', '', '', vStArgument);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.OnArgument (S : String ) ;
begin
  Inherited ;
  FStArgument := S;

  FStGeneral := ReadTokenSt(FStArgument);
  FStDateDeb := ReadTokenSt(FStArgument);
  FStDateFin := ReadTokenSt(FStArgument);

  FStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');

  FListe := THGrid(GetControl('FLISTE', True));

  FListe.ColAligns[cColAgregat]     := TaCenter;
  FListe.ColAligns[cColLibelle]     := TaLeftJustify;
  FListe.ColALigns[cColTotalFiscal] := TaRightJustify;
  FListe.ColAligns[cColEcart]       := TaCenter;
  FListe.OnDblClick                 := OnDblClickFListe;

  FTobRub    := Tob.Create('LESRUBRIQUES', nil, -1);
  FTobFListe := Tob.Create('LAGRILLE', nil, -1);

  FZCumul             := TZCumul.Create;
  FZCumul.ModeFonc    := mfMemory;
  FZCumul.AvecLibelle := True;
  FZCumul.InitCriteres('', '' , '' , '(' + FStDateDeb + ')(' + FStDateFin + ')', '');

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.OnLoad ;
var i,j : integer;
    lTobTemp : Tob;
    LesCptes : TStringList;

    lSoldeComptable : Double;
    LesBalances : TStringList;
    LeResult : TabloExt;

    lDateDeb, lDateFin : TDateTime;

begin
  Inherited ;
  Fliste.BeginUpdate;

  LesBalances := nil;
  try
    LesBalances := TStringList.Create;

    lDateDeb := StrToDate(FStDateDeb);
    lDateFin := StrToDate(FStDateFin);

    lSoldeComptable := GetCumul('GEN', FStGeneral, '', '', '', '', '', lDateDeb, lDateFin, True, True, LesBalances, LeResult , False);

    THPanel(GetControl('PANEL1', True)).Caption := '     ' +
      TraduireMemoire('Compte') + ' : ' + FStGeneral + ' - ' +
      GetColonneSql('GENERAUX', 'G_LIBELLE', 'G_GENERAL = "' + FStGeneral + '"') + ' - solde : ' +
      IIF( lSoldeComptable >= 0 ,
           StrFMontant(lSoldeComptable, 13, V_Pgi.OkDecV, '', True) + ' D',
           StrFMontant(Abs(lSoldeComptable), 13, V_Pgi.OkDecV, '', True) + ' C') +
      ' ' + TraduireMemoire('du') + ' ' + FStDateDeb +
      ' ' + TraduireMemoire('au') + ' ' + FStDateFin;

  finally
    LesBalances.Free;
  end;

  LesCptes := TStringList.Create;
  try
     SelectAgregatDictionnaire('RB_RUBRIQUE, RB_LIBELLE, RB_SIGNERUB, RB_NODOSSIER, RB_PREDEFINI, RB_DATEVALIDITE', '',
                               StrToDate(FStDateFin), FStLiasseDossier, FTobRub);

    for i := 0 to FTobRub.Detail.Count -1 do
    begin
      lTobTemp := FTobRub.Detail[i];
      LesCptes.Clear;
      FZCumul.GetValeur('RUBRIQUE', lTobTemp.GetString('RB_RUBRIQUE'), LesCptes);
      // Parcours des informations des comptes
      for j := 0 to LesCptes.Count - 1 do
      begin
        TraiteInfoCompte(lTobTemp, LesCptes[j]);
      end;
    end;

    if FTobFListe.Detail.Count = 0 then Exit;

    FListe.RowCount := FTobFListe.Detail.Count + 1;
    for i := 0 to FTobFListe.Detail.Count-1 do
    begin
      lTobTemp := FTobFListe.Detail[i];
      FListe.Cells[cColAgregat, i+1]     := lTobTemp.GetString('AGREGAT');
      FListe.Cells[cColLibelle, i+1]     := lTobTemp.GetString('LIBELLE');
      FListe.Cells[cColTotalFiscal, i+1] := lTobTemp.GetString('TOTALFISCAL');
      FListe.Cells[cColEcart, i+1]       := lTobTemp.GetString('ECART');
    end;

  finally
    LesCptes.Free;
    FListe.Row := 1;
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
    FListe.EndUpdate ;
    if FListe.CanFocus then FListe.SetFocus;
  end;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.TraiteInfoCompte(vTobRub : Tob; vSt: string);
var lStTemp    : string;
    lStGeneral : string;
    lDbSM    : Double;
    lDbSD    : Double;
    lDbSC    : Double;
    lDbTD    : Double;
    lDbTC    : Double;
    lDbEcart : Double;
    lInCodification : integer;

    lTobFille : Tob;
begin
  lStTemp := vSt;
  lStGeneral := ReadTokenPipe(lStTemp, ':');

  if lStGeneral <> FStGeneral then Exit;

  lTobFille := Tob.Create('', FTobFliste, -1);
  lTobFille.AddChampSupValeur('AGREGAT', vTobRub.Getstring('RB_RUBRIQUE'));
  lTobFille.AddChampSupValeur('LIBELLE', vTobRub.Getstring('RB_LIBELLE'));

  lDbSM := Valeur( ReadTokenPipe(lStTemp, ':') );           // Sens Mixte
  lDbSC := Valeur( ReadTokenPipe(lStTemp, ':') );           // Sens Créditeur
  lDbSD := Valeur( ReadTokenPipe(lStTemp, ':') );           // Sens Débiteur
  lDbTC := Valeur( ReadTokenPipe(lStTemp, ':') );           // Total Crédit du Compte
  lDbTD := Valeur( ReadTokenPipe(lStTemp, ':') );           // Total Débit du Compte
  lInCodification := StrToInt(ReadTokenPipe(lStTemp, ':')); // Codification

  lDbEcart := 0;
  case lInCodification of
    7 : begin
          lDbEcart := lDbSM;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbSM, 13, V_Pgi.OkDecV, '', True) + ' (SM)');
        end;

    6 : begin
          lDbEcart := lDbSC;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbSC, 13, V_Pgi.OkDecV, '', True) + ' (SC)');
        end;

    5 : begin
          lDbEcart := lDbSD;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbSD, 13, V_Pgi.OkDecV, '', True) + ' (SD)');
        end;

    3 : begin
          lDbEcart := lDbTC;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbTC, 13, V_Pgi.OkDecV, '', True) + ' (TC)');
        end;

    2 : begin
          lDbEcart := lDbTD;
          lTobFille.AddChampSupValeur('TOTALFISCAL', StrFMontant(lDbTD, 13, V_Pgi.OkDecV, '', True) + ' (TD)');
        end;
  end;

  if vTobRub.GetString('RB_SIGNERUB') = 'POS' then
    lDbEcart := lDbTD - lDbTC - lDBEcart
  else
    lDbEcart := lDbTD - lDbTC - (-lDBEcart);

  lDbEcart := Arrondi(lDbEcart, 0);  
  lTobFille.AddChampSupValeur('ECART', IIF(lDBEcart <> 0, TraduireMemoire('Oui'), ''));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/10/2006
Modifié le ... : 05/12/2006    
Description .. : GCO - 05/12/2006 - FQ 19247
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.OnDblClickFListe(Sender: TObject);
var lStTemp  : string;
    lTobTemp : Tob;
begin
  lStTemp := FListe.Cells[cColAgregat, FListe.Row];
  if lStTemp <> '' then
  begin
    lTobTemp := FTobRub.FindFirst(['RB_RUBRIQUE'],[lStTemp], False);
    if lTobTemp <> nil then
      ParametrageRubrique(lTobTemp.GetString('RB_RUBRIQUE') + ';' +
                          lTobTemp.GetString('RB_NODOSSIER') + ';' +
                          lTobTemp.GetString('RB_PREDEFINI') + ';' +
                          DateToStr(lTobTemp.GetDateTime('RB_DATEVALIDITE')), taConsult, CtxRubrique);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTOTALCPTLIASSE.OnClose ;
begin
  FreeAndNil(FZCumul);
  FreeAndNil(FTobRub);
  FreeAndNil(FTobFListe);
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPTOTALCPTLIASSE ] ) ; 
end.
