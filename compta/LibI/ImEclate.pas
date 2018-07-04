{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/07/2003
Modifié le ... : 27/07/2005
Description .. : - 18/07/2003 - Supresssion des champs euro
Suite ........ : - CA - 27/05/2004 - FQ 13199 - Date d'acquisition=date
Suite ........ : d'achat de l'immo de départ et non date d'opération
Suite ........ : - CA - 11/10/2004 - FQ 14789 - Correction bug calcul de
Suite ........ :  l'antériorité en éclatement
Suite ........ : BTY 07/05 FQ 16290 Gérer montants selon nb décimales de la monnaie du dossier
Suite ........ : MBO - 21/11/2005 - FQ 17042 - pb "impossible de focaliser une fenetre..." car on
                 faisait un setfocus sur une zone disable
                 TGA 23/11/2005 la fonction d'enregistrement de l'éclatement devient publique
                 pour être appelée par le changement de méthode
                 enregistreEclatement est dupliqué en New_enregistreEclatement
                 FQ 17291 TGA blocage si mt Ht immo fille avec des décimales
                 FQ 17288 TGA Externalisation du traitement dans imeclate2
                 FQ 17397 TGA 02/02/2006 - Test si opération restante
Suite......... : BTY - 06/06 - FQ 18393 - En série, reprendre la date de l'opération
Suite......... : BTY - 06/06 - FQ 18394 - En série, reprendre le no de serie pour qu'en annulation on puisse annuler les autres de la série
Suite......... : MBO - 10/06 - Modif annulation de l'éclatement pour prise en compte PRIME ou SBV
Suite......... : MVG - 30/10/2006 FQ 19061
Suite......... : BTY 05/07 FQ 19968 Suppression de l'immo éclatée, supprimer l'analytique
*****************************************************************}
unit ImEclate;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ImAncOpe, HSysMenu, hmsgbox, StdCtrls, ComCtrls, HRichEdt, HRichOLE,
  Mask, Hctrls, HTB97, ExtCtrls, HPanel,OpEnCour, db,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Hent1, UTOB, ImEnt,
{$IFNDEF CMPGIS35}
  AMSYNTHESEDPI_TOF,
{$ENDIF}
  LicUtil ;


type
  TFEclatement = class(TFAncOpe)
    HM2: THMsgBox;
    HLabel4: THLabel;
    VALEURECL: THNumEdit;
    HLabel5: THLabel;
    QUANTITE: THNumEdit;
    LIBELLE: TEdit;
    HLabel9: THLabel;
    CODEIMMO: TEdit;
    HLabel6: THLabel;
    BOLDECLATE: TCheckBox;
    procedure VALEURECLExit(Sender: TObject);
    procedure QUANTITEExit(Sender: TObject);
    procedure QUANTITEEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    fQuantiteAvEcl, fValeur : double;
    fNouveauCode : string;
    // mbo - pour conseil de compil car non utilisés fNumPlanAvant, fNumPlanApres : integer;
    fOrdreSerie : integer;  // BTY FQ 18394
    procedure DeLaFiche ;
    procedure EnregistreEclatement;
    function CalcQteEcl : double ;
    function CalcVoEcl : double ;

  public
    { Déclarations publiques }
    procedure InitZones;override;
    function ControleZones : boolean;override;
    procedure New_enregistreEclatement(FFcode:String;FFMttHt:double;FFQuantiteAvEcl:double;
    FFnouveaucode:String;FFValeur:double;FFnew_Valeur:double;FFLibelle:String;FFQuantite:double;
    FFDateope:String;Boldecl:Boolean;FFBlocnote:String );
  end;

// FQ 18393
function ExecuteEclatement(var Code : string) : TModalResult;
//function ExecuteEclatement(Code : string) : TModalResult;
procedure AnnuleEclatement(LogEclatement : TLogEclatement; TabMessErreur : THMsgBox);
procedure TraiteAnnuleLogExceptionnel(CodeOrig,CodeEclate : string;var dMontantExc : double);

const

  HM: array[0..0] of string =
  ('L''immobilisation ne peut pas être éclatée car un élément exceptionnel a été enregistré.');


implementation

uses  Outils,
      ImPlan,
      IMMO_TOM
      {$IFDEF SERIE1}
      , Ut2Points
      , uterreur
      {$ENDIF}
      ,imeclate2;

{$R *.DFM}

function ExecuteEclatement(var Code : string) : TModalResult;
//function ExecuteEclatement(Code : string) : TModalResult;
var  FEclatement: TFEclatement;
     Masque : String ;
     i : integer;

begin
  FEclatement:=TFEclatement.Create(Application) ;
  // FQ 18393
  //FEclatement.fCode:=Code ;
  FEclatement.fCode:= ReadTokenSt(Code);
  FEclatement.fDateSerie := iDate1900;
  if Code <> '' then
     FEclatement.fDateSerie:= StrToDate(ReadTokenSt(Code));
  //
  // FQ 18394
  FEclatement.fOrdreSerie:= TrouveNumeroOrdreSerieLogSuivant;
  if Code <> '' then
     FEclatement.fOrdreSerie:= StrToInt(ReadTokenSt(Code));

  FEclatement.fProcEnreg:=FEclatement.EnregistreEclatement;

  // BTY 07/05 Fiche 16290 :
  // VALEURECL à formater sur nb décimales de la monnaie de saisie du dossier
  Masque:='' ;
  for i:=1 To V_PGI.OkDecV Do  Masque := Masque + '0';
  if V_PGI.OkDecV = 0 then
     Feclatement.ValeurECL.Masks.PositiveMask:='#,##0'
  else
     Feclatement.ValeurECL.Masks.PositiveMask:='#,##0.' ;
  Feclatement.ValeurECL.Masks.PositiveMask := Feclatement.ValeurECL.Masks.PositiveMask + Masque ;
  Feclatement.ValeurECL.Masks.NegativeMask := Feclatement.ValeurECL.Masks.PositiveMask;
  Feclatement.ValeurECL.Masks.ZeroMask := Feclatement.ValeurECL.Masks.PositiveMask;
  //

  try
    // Test si présence d'exceptionnel sur immo à éclater
    // FQ 18393
    //if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+code+'" AND I_MONTANTEXC<>0') then
    if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+FEclatement.fCode+'" AND I_MONTANTEXC<>0') then
      begin
        PGIBox(HM[0]);
        result := mrNo;
      end
    Else
      begin
        FEclatement.ShowModal ;
        result := FEclatement.ModalResult;
        // FQ 18393
        if result = mrYes then
           begin
           Code := FEclatement.DATEOPE.Text;
           // FQ 18394
           Code := Code + ';' + IntToStr(FEclatement.fOrdreSerie);
           end;
      end;
  finally
  FEclatement.Free ;

  end ;
end ;

function TFEclatement.CalcQteEcl : double ;
begin
   if (FValeur=FMttHt) or (FQuantiteAvEcl=1) then result:=FQuantiteAvEcl
                                             else result:=FValeur/(FMttHt/FQuantiteAvEcl);
end ;

function TFEclatement.CalcVoEcl : double ;
begin
  Result:=FQuantite*(fMttHt/FQuantiteAvEcl);
end ;

procedure TFEclatement.InitZones ;
begin
  BOLDECLATE.Visible :=  (V_PGI.PassWord = CryptageSt(DayPass(Date)));
  inherited;
  fQuantiteAvEcl := fQuantite;
  fValeur := fMttHt ;
  fNouveauCode:=NouveauCodeImmo ;
  if FQuantiteAvEcl=1 then
    QUANTITE.Enabled:=false
  else
    begin
    QUANTITE.Min:=1.00;
    QUANTITE.Max:=fQuantite;
    end;
  VALEURECL.Min:=0.00;
  VALEURECL.Max:=fMttHt;
  VALEURECL.Value:=fValeur ;
  QUANTITE.Value:=fQuantite ;
  CODEIMMO.Text:=fNouveauCode ;
  LIBELLE.Text:=fLibelle ;
  // FQ 18393
  if fDateSerie <> iDate1900 then
     DATEOPE.Text := DateToStr(fDateSerie);
end ;

procedure TFEclatement.DeLaFiche ;
begin
  fValeur:=VALEURECL.Value ;
  if QUANTITE.Text<>'' then  fQuantite:=StrToFloat(QUANTITE.Text)
  else fQuantite:=0.00;
  fNouveauCode:=CODEIMMO.Text ;
  fLibelle:=LIBELLE.Text ;
end ;

function TFEclatement.ControleZones : boolean;
{$IFDEF SERIE1}
var
  Err : TPGIErr ; //XMG 24/10/02
{$ENDIF}
begin
  result:=False ;
  if (inherited ControleZones) then
  begin
    Result := True;
    DeLaFiche;
    // Eclatement Total
    if (FMttHt=FValeur) and (FQuantiteAvEcl=FQuantite) then
    begin
      HM2.execute(2,Caption,'');
      ModalResult := mrNone;
      result := false;
    end else
    if (fQuantite > FQuantiteAvEcl) then   // mbo fq 17042 or (FValeur>FMttHt)  then
    begin
      HM2.execute(4,Caption,'');
      ModalResult := mrNone;
      QUANTITE.SetFocus;
      result:=false;
    end else if (FValeur>FMttHt) then  // mbo fq 17042 eclater le test en 2
    begin
      HM2.execute(4,Caption,'');
      ModalResult := mrNone;
      VALEURECL.SetFocus;
      result:=false;
    end else if (fQuantite > Int(fQuantite)) then
    begin
      HM2.execute(3,Caption,'');
      ModalResult := mrNone;
      QUANTITE.SetFocus;
      result:=false;
    end else if ((fValeur=0.00) or (fQuantite=0.00) or (fNouveauCode='') or (fLibelle='')) then
    begin
      if fValeur=0.00 then VALEURECL.SetFocus
      else if fQuantite=0.00 then QUANTITE.SetFocus
      else if fNouveauCode='' then CODEIMMO.SetFocus
      else if fLibelle='' then LIBELLE.SetFocus;
      HM2.execute(0,Caption,'');
      ModalResult := mrNone;
      result:=false;
    end
{$IFDEF SERIE1}
    else  if not Testecode(fNouveauCode,Err) then
    begin
      PGIBox(Err.Libelle,Caption) ;
      FocusControle(CODEIMMO) ;
      ModalResult:=mrNone ;
      Result:=FALSE ;
    end
{$ENDIF}
    else if ExisteSQL('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+fNouveauCode+'"') then
    begin
      HM2.execute(1,Caption,'');
      FocusControl(CODEIMMO);
      ModalResult := mrNone;
      result:=false;
    end;
  end;
end;

procedure TFEclatement.VALEURECLExit(Sender: TObject);
begin
  inherited;
  fValeur:=VALEURECL.Value;
  fQuantite:=CalcQteEcl;
  QUANTITE.Value:=fQuantite ;
end;

procedure TFEclatement.QUANTITEExit(Sender: TObject);
begin
  inherited;
  if QUANTITE.Text<>'' then  fQuantite:=StrToFloat(QUANTITE.Text)
                       else fQuantite:=0.00;
  QUANTITE.Value:=fQuantite ;
  if fQuantiteAvEcl<>1 then  fValeur:=CalcVoEcl
                       else fValeur:=VALEURECL.Value;
  VALEURECL.Value:=fValeur ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/10/2004
Modifié le ... :   /  /
Description .. : Appel de l'éclatement
Mots clefs ... :
*****************************************************************}
procedure TFEclatement.EnregistreEclatement;
var FFnew_Valeur : Double;
    Boldecl : Boolean;
    FFDateope : string;
    FFBlocnote :String;
begin

   {==============================================================
    TGA 23/11/2005 Appel nouvelle fonction
           FFcode = Fcode = code immo mère
           FFMttHt = FMttHt = montant ht immo mère
           FFQuantiteAvEcl = FQuantiteAvEcl = Quantité immo mère
           FFnouveaucode = Fnouveaucode = code immo fille
           FFValeur = FValeur = montant ht immo fille
           FFnew_valeur = nouvelle valorisation immo fille
           FFLibelle = FLibelle = libelle immo fille
           FFQuantite = FQuantite = Quantité immo fille
           FFDateOpe = DateOpe = Date opération éclatement
           Boldecl = Boldeclate = type de traitement
   ==============================================================}
   FFnew_Valeur:=0;
   FFDateope := DATEOPE.text;
   FFBlocnote:= RichToString (Il_Blocnote);
   if BOLDECLATE.Checked then
      Boldecl:=True
   else
      Boldecl:=False;

   //MVG FQ 19061
   enregistreEclatement_2(Fcode,FMttHt,FQuantiteAvEcl,Fnouveaucode,
   Arrondi(FValeur,V_PGI.OkDecV),
   FFnew_Valeur,FLibelle,FQuantite,FFDateope,Boldecl,FFBlocnote,
   fOrdreSerie);  // FQ 18394

   //==============================================================
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/10/2004
Modifié le ... : 12/10/2006
Description .. : Annulation de l'éclatement
Suite ........ : modif mbo  11.10.2006 = pour annulation d'un éclatement 
Suite ........ : avec présence d'une prime ou d'une subvention sur la mère
Mots clefs ... : 
*****************************************************************}
procedure AnnuleEclatement(LogEclatement : TLogEclatement; TabMessErreur : THMsgBox);

var T1,T2,T3,T4: TOB ; i: integer ; OkEche: boolean ;  Q1,Q2: TQuery ;
   prorata:double;
   valorisation:boolean;
   Qlog : TQuery ;
   x: Integer;
   prorat : Double;
   GestPRI : boolean;
   GestSBV : boolean;
   SAVreprisePRI :double;
   SAVrepriseSBV :double;
   SAVmntPRI :double;
   SAVmntSBV :double;


begin
  prorata:=0;
  valorisation :=False;

  // ajout mbo pour prime et sbv
  SAVreprisePRI := 0.00;
  SAVrepriseSBV := 0.00;
  SAVmntPRI := 0.00;
  SAVmntSBV := 0.00;



  if IsOpeEnCours(nil,LogEclatement.CodeEclate,false) then
    begin
    TabMessErreur.execute(21,'','');
    V_PGI.IoError:=oeUnknown ;
    raise Exception.Create (''); //YCP ??
    exit;
    end;

  // TGA 30/11/2005 lecture d'immolog pour déterminer si l'immo à été
  // éclatée et revalorisée.
  // TGA 02/02/2006 ajout test sur ordre

  // modif pour mbo
  //QLog:=OpenSQL('SELECT IL_MONTANTDOT,IL_CODECB FROM IMMOLOG WHERE IL_IMMO="'+LogEclatement.lequel+
  //      '" AND IL_TYPEOP ="ECL" AND IL_MONTANTDOT<>0 AND IL_ORDRE="'+IntToStr(LogEclatement.Ordre)+'"', FALSE) ;

  QLog:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+LogEclatement.lequel+
        '" AND IL_TYPEOP ="ECL" AND IL_ORDRE="'+IntToStr(LogEclatement.Ordre)+'"', FALSE) ;

  if not Qlog.Eof then
  Begin
    if Qlog.FindField('IL_MONTANTDOT').AsFloat <> 0 then
    begin
       prorat := Qlog.FindField('IL_MONTANTDOT').AsFloat;
       // TGA 20/01/2006 test si sauvegarde > 1 ou pas (voir Imeclate2)
       IF prorat > 1 Then
         Begin
           x:=length(strfpoint(INT(Prorat)));
           For i:=1 To x DO Prorat:=(Prorat/10) ;
         end;
       prorata := (1/(1-Prorat));
       valorisation :=True;
    end;

    //ajout mbo pour gestion prime et subvention
    SAVreprisePRI := Qlog.FindField('IL_REVISIONREPECO').AsFloat;
    SAVrepriseSBV := Qlog.FindField('IL_REVISIONREPFISC').AsFloat;
    SAVmntPRI := Qlog.FindField('IL_REVISIONDOTECO').AsFloat;
    SAVmntSBV := Qlog.FindField('IL_REVISIONDOTFISC').AsFloat;
  End;
  Ferme(Qlog) ;

  T1:=Tob.Create('IMMO',nil,-1) ; T1.SelectDB('"'+LogEclatement.Lequel+'"',nil) ;
  T2:=Tob.Create('IMMO',nil,-1) ; T2.SelectDB('"'+LogEclatement.CodeEclate+'"',nil) ;
  OkEche:=(T1.GetValue('I_NATUREIMMO')='CB') or (T1.GetValue('I_NATUREIMMO')='LOC')  ;

  // ajout mbo pour gestion prime et subvention
  GestPRI := (T1.GetValue('I_SBVPRI')<> 0);   // on regarde si gestion prime ou sbv sur la mère
  GestSBV := (T1.GetValue('I_SBVMT')<> 0);

  // Réinitialisation de l'immo mère
  IF valorisation =True THEN
    Begin
      for i:=low(LesChamps1) to high(LesChamps1) do CalculSurChampsTob(T1,LesChamps1[i],'*',prorata,V_PGI.OkDecV) ;

      if OkEche then
        for i:=low(LesChamps2) to high(LesChamps2) do CalculSurChampsTob(T1,LesChamps2[i],'*',prorata,V_PGI.OkDecV) ;

      //ajout mbo pour prime et sbv
      if GestPRI then
        for i:=low(LesChamps3) to high(LesChamps3) do CalculSurChampsTob(T1,LesChamps3[i],'*',prorata,V_PGI.OkDecV) ;

      if GestSBV then
        for i:=low(LesChamps4) to high(LesChamps4) do CalculSurChampsTob(T1,LesChamps4[i],'*',prorata,V_PGI.OkDecV) ;
    end
  ELSE
    begin
      for i:=low(LesChamps1) to high(LesChamps1) do CalculSurChampsTob(T1,LesChamps1[i],'+',T2.GetValue(LesChamps1[i]),V_PGI.OkDecV) ;

      if OkEche then
         for i:=low(LesChamps2) to high(LesChamps2) do CalculSurChampsTob(T1,LesChamps2[i],'+',T2.GetValue(LesChamps2[i]),V_PGI.OkDecV) ;

      // ajout mbo pour prime et subvention
      // On doit faire une gestion manuelle car on a perdu les montant prime et sbv de la fille qui
      // sont supprimées avant d'annuler l'éclatement
      if GestPRI then
      begin
         T1.PutValue('I_REPRISEUO', SAVreprisePRI);
         T1.PutValue('I_SBVPRI', SAVmntPri);
      end;
      if GestSBV then
      begin
         T1.PutValue('I_CORRECTIONVR', SAVrepriseSBV);
         T1.PutValue('I_SBVMT', SAVmntSBV);
      end;
    end;


  //if LogEclatement.QteEclate>=1 then T1.PutValue('I_QUANTITE',T1.GetValue('I_QUANTITE')+T2.GetValue('I_QUANTITE')) ;
  IF (T1.GetValue('I_QUANTITE')=1) and (T2.GetValue('I_QUANTITE')=1) Then
     T1.PutValue('I_QUANTITE',1)
  Else
     T1.PutValue('I_QUANTITE',T1.GetValue('I_QUANTITE')+T2.GetValue('I_QUANTITE')) ;

  T1.PutValue('I_CHANGECODE','') ;

  // suppression dans immomaor du plan eclatement de la mère
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+LogEclatement.Lequel+'" AND IA_NUMEROSEQ='+IntToStr(T1.GetValue('I_PLANACTIF')));
  T1.PutValue('I_PLANACTIF',LogEclatement.PlanActifAv) ;

  // TGA 02/02/2006 - Suppression de l'immo fille issue de l'éclatement
  // remontée de opencours sinon test de présence avant suppression
  ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+LogEclatement.Lequel+
             '" AND IL_ORDRE="'+IntToStr(LogEclatement.Ordre)+'"') ;

  // TGA 02/02/2006 - Test s'il reste un éclatement sur l'immo mère
  if not ExisteSQL('SELECT IL_IMMO FROM IMMOLOG WHERE (IL_IMMO="'+LogEclatement.Lequel+
    '") AND (IL_TYPEOP = "ECL") AND (IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'") AND (IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+'")') then
   T1.PutValue('I_OPEECLATEMENT','-');

  // TGA 02/02/2006 - FQ 17397 Test si opération restante <> d'acquisition ou clôture
  if not ExisteSQL('SELECT IL_IMMO FROM IMMOLOG WHERE (IL_IMMO="'+LogEclatement.Lequel+
    '") AND (IL_TYPEOP NOT IN ("ACQ","CLO")) AND (IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'") AND (IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+'")') then
     T1.PutValue('I_OPERATION','-');

  T1.InsertOrUpdateDB ;
  T1.Free ;
  T2.Free ;

  //Tga 28/06/2006 Maj Immomvtd suppression de l'éclatement
{$IFNDEF CMPGIS35}
  AM_MAJ_IMMOMVTD('W',LogEclatement.lequel,LogEclatement.CodeEclate,0);
{$ENDIF}


  //ImmoEche
  if OkEche then
    begin
    Q1:=OpenSQL('SELECT * FROM IMMOECHE WHERE IH_IMMO="'+LogEclatement.Lequel+'" ORDER BY IH_DATE',false);
    T1:=Tob.Create('IMMOECHE',nil,-1) ; T1.LoadDetailDB('IMMOECHE','','',Q1,true) ;
    Ferme(Q1) ;
    Q2:=OpenSQL('SELECT * FROM IMMOECHE WHERE IH_IMMO="'+LogEclatement.CodeEclate+'" ORDER BY IH_DATE', FALSE);
    T2:=Tob.Create('IMMOECHE',nil,-1) ; T2.LoadDetailDB('IMMOECHE','','',Q2,true) ;
    Ferme(Q2) ;
    for i:=0 to T1.Detail.Count-1 do
      begin
      T3:=T1.Detail[i] ;
      T4:=T2.Detail[i] ;
      // TGA 30/11/2005
      IF valorisation =True THEN
        Begin
         T3.PutValue('IH_MONTANT',T3.GetValue('IH_MONTANT')*prorata) ;
         T3.PutValue('IH_FRAIS',T3.GetValue('IH_FRAIS')*prorata) ;
        End
      else
        Begin
         T3.PutValue('IH_MONTANT',T3.GetValue('IH_MONTANT')+T4.GetValue('IH_MONTANT')) ;
         T3.PutValue('IH_FRAIS',T3.GetValue('IH_FRAIS')+T4.GetValue('IH_FRAIS')) ;
        End;
      end ;
    T1.UpdateDB ;
    T1.Free ;
    T2.Free ;
    end ;

  // Plan d'unités d'oeuvre
  Q1:=OpenSQL('SELECT * FROM IMMOUO WHERE IUO_IMMO="'+LogEclatement.Lequel+'" ORDER BY IUO_DATE',false);
  T1:=Tob.Create('IMMOUO',nil,-1) ;
  T1.LoadDetailDB('IMMOUO','','',Q1,true) ;
  Ferme(Q1) ;
  Q2:=OpenSQL('SELECT * FROM IMMOUO WHERE IUO_IMMO="'+LogEclatement.CodeEclate+'" ORDER BY IUO_DATE', FALSE);
  T2:=Tob.Create('IMMOUO',nil,-1) ; T2.LoadDetailDB('IMMOUO','','',Q2,true) ;
  Ferme(Q2) ;
  for i:=0 to T1.Detail.Count-1 do
  begin
    T3:=T1.Detail[i] ;
    T4:=T2.Detail[i] ;
    // TGA 30/11/2005
    IF valorisation =True THEN
      T3.PutValue('IUO_UNITEOEUVRE',T3.GetValue('IUO_UNITEOEUVRE')*prorata)
    else
      T3.PutValue('IUO_UNITEOEUVRE',T3.GetValue('IUO_UNITEOEUVRE')+T4.GetValue('IUO_UNITEOEUVRE')) ;
  end ;
  T1.UpdateDB ;
  T1.Free ;
  T2.Free ;

  // New Be One cannot be
  // Suppression de l'immobilisation résultant de l'éclatement
  ExecuteSQL('DELETE FROM IMMO WHERE I_IMMO="'+LogEclatement.CodeEclate+'"') ;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+LogEclatement.CodeEclate+'"') ;
  ExecuteSQL('DELETE FROM IMMOECHE WHERE IH_IMMO="'+LogEclatement.CodeEclate+'"') ;
  ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+LogEclatement.CodeEclate+'"') ;
  ExecuteSQL('DELETE FROM IMMOUO WHERE IUO_IMMO="'+LogEclatement.CodeEclate+'"') ;
  // FQ 19968 suppression des ventilations analytiques générées
  ExecuteSQL('DELETE FROM VENTIL WHERE V_COMPTE="'+LogEclatement.CodeEclate+'"') ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/10/2004
Modifié le ... :   /  /
Description .. : Annulation
Mots clefs ... :
*****************************************************************}
procedure TraiteAnnuleLogExceptionnel(CodeOrig,CodeEclate : string;var dMontantExc : double);
var
  ListeLog : TList;
  EnrListeLog : PLogChPlan;
  sTypeExc,sMontantDot,Requete : string;
  Ind,iPos : integer;
  dMttExc : double;
  QLog : TQuery;
begin
  QLog:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+CodeEclate+'" AND IL_TYPEOP="EEC" ORDER BY IL_DATEOP ASC', FALSE) ;
  if not QLog.EOF then QLog.First;
  ListeLog := TList.Create;
  New(EnrListeLog);
  while not QLog.EOF do
  begin
    QLog.Edit;
    TraiteEnregLogChPlan(QLog,EnrListeLog^);
    EnrListeLog^.TypeExc:=QLog.FindField('IL_TYPEDOT').AsString;
    EnrListeLog^.MontantExc:=QLog.FindField('IL_MONTANTDOT').AsFloat;
    ListeLog.Add(EnrListeLog);
    QLog.Next;
  end;
  Ferme(QLog);
  dMontantExc:=0.00;dMttExc:=0.00;sTypeExc:='';
  for Ind := 0 to (ListeLog.Count - 1) do
  begin
    sMontantDot:=FloatToStrF(EnrListeLog^.MontantExc,ffFixed,20,V_PGI.OkDecV);
    iPos := Pos(',', sMontantDot);
    sMontantDot[iPos]:='.';
    Requete:='SELECT * FROM IMMOLOG WHERE (IL_IMMO="'+CodeOrig+'"'+
             ' AND IL_DATEOP="'+USDateTime(EnrListeLog^.DateOpe)+'"'+
             ' AND IL_MONTANTDOT='+sMontantDot+
             ' AND IL_TYPEOP="EEC")';
    QLog:=OpenSQL(Requete, FALSE) ;
    QLog.Edit ;
    EnrListeLog := ListeLog.Items[Ind];
    dMontantExc := dMontantExc + EnrListeLog^.MontantExc;
    QLog.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', 'ELE', FALSE);
    QLog.FindField('IL_TYPEOP').AsString := 'ELE';
    QLog.FindField('IL_MONTANTEXC').AsFloat :=
      QLog.FindField('IL_MONTANTEXC').AsFloat-dMttExc;
    QLog.FindField('IL_MONTANTDOT').AsFloat :=
        QLog.FindField('IL_MONTANTDOT').AsFloat+EnrListeLog^.MontantExc;
    QLog.Post;
    Ferme(QLog);
    dMttExc:=EnrListeLog^.MontantExc;
  end;
  ListeLog.Free
end;

procedure TFEclatement.QUANTITEEnter(Sender: TObject);
begin
  inherited;
  fQuantite := 0.0;
end;

procedure TFEclatement.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
HelpContext:=511030 ;
{$ELSE}
HelpContext:=2111200 ;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Thierry GAUZAN
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. : Traitement pour appel provisoire par  changement de méthode
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TFEclatement.New_enregistreEclatement(FFcode: String;
                                                FFMttHt : double;
                                                FFQuantiteAvEcl : double;
                                                FFnouveaucode: String;
                                                FFValeur : double;
                                                FFnew_Valeur : double;
                                                FFLibelle : String;
                                                FFQuantite : double;
                                                FFDateope : String;
                                                Boldecl : Boolean;
                                                FFblocnote:String);

var FFOrdreSerie : integer;
begin
   // FQ 18394 Passer le n° d'ordre en série
   FFOrdreSerie :=  TrouveNumeroOrdreSerieLogSuivant;

   // MVG FQ 19061
   enregistreEclatement_2(FFcode,FFMttHt,FFQuantiteAvEcl,FFnouveaucode,
   Arrondi(FFValeur,V_PGI.OkDecV),
   FFnew_Valeur,FFLibelle,FFQuantite,FFDateope,Boldecl,FFBlocnote,
   FFOrdreSerie);  // FQ 18394

end;

end.


