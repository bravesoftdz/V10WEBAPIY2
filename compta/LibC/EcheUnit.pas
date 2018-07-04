{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit EcheUnit;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
//  Graphics,
  Controls,   // mrOk, mrCancel
  Forms,
//  Dialogs,
  StdCtrls,   // TLabel...
  Hctrls,     // THLabel, THValComboBox, THNumEdit, AddMenuPop, RechDom, OpenSQL, Ferme, USDateTime, ExecuteSQL
  Mask,       // TMaskEdit
  Buttons,    // TBitBtn
  ExtCtrls,   // TPanel
  Ent1,       // VH, ChangeMask, AfficheLeSolde, OkSynchro, LienS1S3
  HEnt1,      // String3, SyncrDefault, V_PGI, tsGene, Arrondi
  Saisutil,   // RDEVISE, RMVT, GetInfoDevise
  hmsgbox,    // THMsgBox, MessageAlerte
  SaisComm,   // WhereEcriture, EstCollFact
{$IFDEF EAGLCLIENT}
  UTOB,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}// TQuery, Eof, Fields
{$ENDIF}
  HSysMenu,   // THSystemMenu
  UtilPGI     // DeviseToPivot, DevieToEuro, PivotToEuro, EuroToPivot
  ;

Type T_ECHEUNIT = RECORD
                  DateEche,DateComptable,DateModif : TDateTime ;
                  ModePaie,Devise  : String3 ;
                  Debit,Credit : Double ;
                  DebitDEV,CreditDEV   : Double ;
                  //DebitEuro,CreditEuro : Double ;
                  TauxDEV              : Double ;
                  //SaisieEuro           : boolean ; 
                  TabTva               : Array[1..5] of Double ;
                  ModeSaisie           : String ;
                  END ;

type
  TFEcheUnit = class(TForm)
    PanelBouton: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    Panel1  : TPanel;
    Panel2  : TPanel;
    HLabel1 : THLabel;
    HLabel2 : THLabel;
    HLabel3 : THLabel;
    FDateEche: TMaskEdit;
    FModePaie: THValComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    E_JOURNAL: TLabel;
    E_NATUREPIECE: TLabel;
    E_DATECOMPTABLE: TLabel;
    E_NUMEROPIECE: TLabel;
    FMontant: THNumEdit;
    HLabel4: THLabel;
    HLabel5: THLabel;
    FDateEche2: TMaskEdit;
    FModepaie2: THValComboBox;
    FTitreSup: TLabel;
    HME: THMsgBox;
    HLabel6: THLabel;
    MTANTORIG: THNumEdit;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
  private
    DEV  : RDEVISE ;
    NowFutur : TDateTime ;
    EcheSup,SensDebit,BQE : Boolean ;
    ResteD,ResteP    : Double ;
    procedure ModifieEche ;
    procedure SauvegardeEche ;
    procedure NewEche ;
    function  PlusDePlace : boolean ;
  public
    M  : RMVT ;
    EU : T_ECHEUNIT ;
  end;

function ModifUneEcheance ( M : RMVT ; EU : T_ECHEUNIT ) : boolean ;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}
  ULibTrSynchro;

function ModifUneEcheance ( M : RMVT ; EU : T_ECHEUNIT ) : boolean ;
Var X : TFEcheUnit ;
    ii  : integer ;
BEGIN
  X:=TFEcheUnit.Create(Application) ;
  Try
    X.M:=M ; X.EU:=EU ;
    ii:=X.ShowModal ;
    Result:=(ii=mrOk) ;
  Finally
    X.Free ;
  End ;
  Screen.Cursor:=SyncrDefault ;
END ;

procedure TFEcheUnit.FormCreate(Sender: TObject);
begin
  EcheSup:=False ;
  BQE:=False ;
  PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFEcheUnit.FormShow(Sender: TObject);
Var Q : TQuery ;
begin
  FDateEche.Text:=DateToStr(EU.DateEche) ;
  FDateEche2.Text:=DateToStr(EU.DateEche) ;
  FModePaie.Value:=EU.ModePaie ;
  FModePaie2.Value:=EU.ModePaie ;
  DEV.Code:=EU.Devise ;
  GetInfosDevise(DEV) ;

  ChangeMask(FMontant,DEV.Decimale,DEV.Symbole) ;
  ChangeMask(MtantOrig,DEV.Decimale,DEV.Symbole) ;
  if EU.Devise<>V_PGI.DevisePivot then BEGIN
    AfficheLeSolde(FMontant,EU.DebitDEV,EU.CreditDEV) ;
    AfficheLeSolde(MtantOrig,EU.DebitDEV,EU.CreditDEV) ;
    END
  else BEGIN
     AfficheLeSolde(FMontant,EU.Debit,EU.Credit) ;
     AfficheLeSolde(MtantOrig,EU.Debit,EU.Credit) ;
   END ;

  E_JOURNAL.Caption:=RechDom('ttJournaux',M.Jal,False) ;
  E_NATUREPIECE.Caption:=RechDom('ttNaturePiece',M.Nature,False) ;
  E_NUMEROPIECE.Caption:=IntToStr(M.Num) ;
  E_DATECOMPTABLE.Caption:=DateToStr(M.DateC) ;
  SensDebit:=(EU.DebitDev<>0) ;

  Q:=OpenSQL('Select J_NATUREJAL from JOURNAL Where J_JOURNAL="'+M.Jal+'"',True) ;
  if Not Q.EOF then BQE:=((Q.Fields[0].AsString='BQE') or (Q.Fields[0].AsString='CAI')) ;
  Ferme(Q) ;

  if BQE then FMontant.Enabled:=False ;
  If (EU.ModeSaisie='BOR') Or (EU.ModeSaisie='LIB') Then FMontant.Enabled:=False ;
end;

procedure TFEcheUnit.NewEche ;
var
    NumEche,k : integer;

    Source : TQuery;
    Dest : TOBM;

    Solde : double;
    Collectif : String;
begin
    Source := OpenSQL('SELECT * FROM ECRITURE WHERE '+WhereEcriture(tsGene,M,False) + ' AND E_NUMLIGNE='+IntToStr(M.NumLigne)+' ORDER BY E_NUMECHE DESC',false);

    if (Not Source.EOF) then
    begin
        Dest := TOBM.Create(EcrGen,'',false);
        Dest.ChargeMvt(Source);

        NumEche := Dest.GetMvt('E_NUMECHE') + 1;
        Dest.PutMvt('E_DATEECHEANCE',StrToDate(FDateEche2.Text));
        Dest.PutMvt('E_MODEPAIE',FModePaie2.Value);
        Dest.PutMvt('E_NUMECHE',NumEche);
        Dest.PutMvt('E_DATEMODIF',NowFutur);
        Dest.PutMvt('E_UTILISATEUR',V_PGI.User);
        Dest.PutMvt('E_DATECREATION',Date);
        if (SensDebit) then
        begin
            Dest.PutMvt('E_DEBITDEV',ResteD);
            Dest.PutMvt('E_DEBIT',ResteP);
        end
        else
        begin
            Dest.PutMvt('E_CREDITDEV',ResteD);
            Dest.PutMvt('E_CREDIT',ResteP);
        end;
        {#TVAENC}
        if VH^.OuiTvaEnc then
        begin
            Collectif := Dest.GetMvt('E_GENERAL');
            if EstCollFact(Collectif) then
            begin
                Solde := EU.Debit + EU.Credit;
                if Solde <> 0 then
                begin
                    for k:=1 to 4 do Dest.PutMvt('E_ECHEENC'+IntToStr(k),Arrondi(EU.TabTva[k]*ResteP/Solde,V_PGI.OkDecV)) ;
                    Dest.PutMvt('E_ECHEDEBIT',Arrondi(EU.TabTva[5]*ResteP/Solde,V_PGI.OkDecV)) ;
                end;
            end;

        end;
        Dest.SetCotation(0) ;
        Dest.SetMPACC;
        {JP 12/04/05 : FQ 15608 : gestion de la synchronisation avec la Trésorerie}
        if EstComptaTreso then
          MajTresoEcritureTOB(Dest, taCreat);

        Dest.InsertDB(nil);
    end;
end;
// Fin BPY

procedure TFEcheUnit.SauvegardeEche ;
Var MP,MD,Coef : Double ;
    SQL   : String ;
    k     : integer ;
    StSup : String ;
BEGIN
StSup:='' ;
StSup:=', E_PAQUETREVISION=1 ' ;
if Not EcheSup then
   BEGIN
   // GG COM
   SQL:='UPDATE ECRITURE SET E_MODEPAIE="'+FModePaie.Value+'", '
       +'E_CODEACCEPT="'+MPTOACC(FModePaie.Value)+'", '  // FQ 15821 - CA - 09/06/2005
       +'E_DATEECHEANCE="'+UsDateTime(StrToDate(FDateEche.Text))+'", E_DATEMODIF="'+UsTime(NowFutur)+'" '
       {JP 12/04/05 : FQ 15608 : Gestion de la synchronisation en Treéosrerie}
       + MajETresoSynchro(M)
       // GG COM
       +StSup
       +'Where '+WhereEcriture(tsGene,M,True)+' AND E_DATEMODIF="'+UsTime(EU.DateModif)+'"' ;
   {JP 12/04/05 : FQ 15608 : Gestion de la synchronisation en Treéosrerie}
   if ExecuteSQL(SQL)<>1 then V_PGI.IoError:=oeUnknown ;
   END
   ELSE
   BEGIN
     if EU.Devise<>V_PGI.DevisePivot then
     BEGIN
       MD:=FMontant.Value ;
       //ME:=DeviseToPivot(MD,EU.TauxDEV,DEV.Quotite) ;
       MP:=DeviseToEuro(MD,EU.TauxDEV,DEV.Quotite) ;
       {JP 01/06/07 : FQ 20443 : le plantage étant aléatoire, peut-être vient-il d'un problème d'arrondi
                      comme pour la FQ 20332}
       if Arrondi((EU.DebitDev+EU.CreditDev), V_PGI.OkDecV)<>0 then
         Coef := MD/(EU.DebitDev+EU.CreditDev)
       else
         Coef := 0 ;
     END
     ELSE
     BEGIN
       MP:=FMontant.Value ; MD:=MP ;
       //ME:=EuroToPivot(MP);
       {JP 01/06/07 : FQ 20443 : le plantage étant aléatoire, peut-être vient-il d'un problème d'arrondi
                      comme pour la FQ 20332}
       if Arrondi((EU.Debit+EU.Credit), V_PGI.OkDecV) <>0 then
         Coef := MP/(EU.Debit+EU.Credit)
       else Coef := 0 ;
     END ;

   if SensDebit then
      BEGIN
      SQL:='UPDATE ECRITURE SET E_MODEPAIE="'+FModePaie.Value+'", '
           +'E_CODEACCEPT="'+MPTOACC(FModePaie.Value)+'", '  // FQ 15821 - CA - 09/06/2005
           +'E_DATEECHEANCE="'+UsDateTime(StrToDate(FDateEche.Text))+'", E_DATEMODIF="'+UsTime(NowFutur)+'", '
           +'E_DEBITDEV='+StrFPoint(MD)+', E_DEBIT='+StrFPoint(MP)
           +', E_CREDITDEV=0, E_CREDIT=0 '
           {JP 12/04/05 : FQ 15608 : Gestion de la synchronisation en Treéosrerie}
           + MajETresoSynchro(M)
           // GG COM
           +StSup
      END else
      BEGIN
      SQL:='UPDATE ECRITURE SET E_MODEPAIE="'+FModePaie.Value+'", '
           +'E_CODEACCEPT="'+MPTOACC(FModePaie.Value)+'", '  // FQ 15821 - CA - 09/06/2005
           +'E_DATEECHEANCE="'+UsDateTime(StrToDate(FDateEche.Text))+'", E_DATEMODIF="'+UsTime(NowFutur)+'", '
           +'E_CREDITDEV='+StrFPoint(MD)+', E_CREDIT='+StrFPoint(MP)
           +', E_DEBITDEV=0, E_DEBIT=0 '
           {JP 12/04/05 : FQ 15608 : Gestion de la synchronisation en Treéosrerie}
           + MajETresoSynchro(M)
           // GG COM
           +StSup ;
      END ;
   {#TVAENC}
    if ((VH^.OuiTvaEnc) and (Coef<>0)) then BEGIN
      for k:=1 to 4 do
        SQL:=SQL+', E_ECHEENC'+IntToStr(k)+'='+StrfPoint(Arrondi(EU.TabTva[k]*Coef,V_PGI.OkDecV)) ;
      SQL:=SQL+', E_ECHEDEBIT='+StrfPoint(Arrondi(EU.TabTva[5]*Coef,V_PGI.OkDecV)) ;
    END ;
    SQL:=SQL+' Where '+WhereEcriture(tsGene,M,True)+' AND E_DATEMODIF="'+UsTime(EU.DateModif)+'"' ;
    {JP 12/04/05 : FQ 15608 : Gestion de la synchronisation en Treéosrerie}
    if ExecuteSQL(SQL)<>1 then V_PGI.IoError:=oeUnknown ;
    if V_PGI.Ioerror=oeOk then NewEche ;
  END ;
END ;

procedure TFEcheUnit.ModifieEche ;
BEGIN
if Transactions(SauvegardeEche,5)<>oeOK then BEGIN
  MessageAlerte(HME.Mess[1]) ;  // Modification non effectuée.
  ModalResult:=mrCancel ;
END ;
END ;

function TFEcheUnit.PlusDePlace : boolean ;
Var Q : TQuery ;
BEGIN
  Result:=False ;
  Q:=OpenSQL('Select Max(E_NUMECHE) from ECRITURE Where '+WhereEcriture(tsGene,M,False)
            +' AND E_NUMLIGNE='+IntToStr(M.NumLigne),True) ;
  if Not Q.EOF then BEGIN
    if Q.Fields[0].AsInteger>=12 then BEGIN
      HME.Execute(5,'','') ;  // Une écriture ne peut pas avoir plus de douze échéances.
      Result:=True ;
    END ;
  END ;
  Ferme(Q) ;
END ;

procedure TFEcheUnit.BValiderClick(Sender: TObject);
Var XD,XP : Double ;
    DDE   : TDateTime ;
    Equil : Boolean ;
    Dec   : Integer; {JP 15/05/07 : FQ 20332}
begin
NowFutur:=NowH ;
  if Not EcheSup then BEGIN
    if FMontant.Value=0 then BEGIN
      HME.Execute(2,'','') ;  // Vous ne pouvez pas renseigner un montant nul.
      Exit ;
    END ;
    if Not IsValidDate(FDateEche.Text) then BEGIN
      HME.Execute(3,'','') ;  // Vous devez renseigner une date valide.
      Exit ;
    END ;
   DDE:=StrToDate(FDateEche.Text) ;
    if Not NbJoursOK(EU.DateComptable,DDE) then BEGIN
      HME.Execute(4,'','') ;  // La date d'échéance doit respecter la plage de saisie autorisée.
      Exit ;
    END ;
    if EU.Devise<>V_PGI.DevisePivot then BEGIN
      if Abs(FMontant.Value)>Abs(EU.DebitDev+EU.CreditDEV) then BEGIN
        HME.Execute(0,'','') ;  // Vous ne pouvez pas modifier le montant à la hausse.
        Exit ;
      END ;
      Equil:=Arrondi(FMontant.Value-(EU.DebitDev+EU.CreditDEV),DEV.Decimale)=0 ;
      END
    else BEGIN
      {JP 15/05/07 : FQ 20332 : gestion des arrondis pour éviter le déclenchement du message 0}
      Dec := DEV.Decimale;
      if Dec <= 0 then Dec := V_PGI.OkDecV;
      
      if Arrondi(Abs(FMontant.Value), Dec) > Arrondi(Abs(EU.Debit+EU.Credit), Dec) then begin
        HME.Execute(0,'','') ;  // Vous ne pouvez pas modifier le montant à la hausse.
        Exit;
      end;
      Equil := Arrondi(FMontant.Value - (EU.Debit + EU.Credit),Dec) = 0;
    end;

    if Equil then BEGIN
      ModifieEche ;
      if V_PGI.IOError=oeOk then ModalResult:=mrOK ;
      Exit ;
      END
    else BEGIN
      if ((LienS1) or (PlusDePlace)) then Exit ;
    END ;

   EcheSup:=True ;
   XD:=EU.DebitDEV+EU.CreditDEV ;
   XP:=EU.Debit+EU.Credit ;
   //XE:=EU.DebitEuro+EU.CreditEuro ;

   if EU.Devise<>V_PGI.DevisePivot then
      BEGIN
        ResteD:=Arrondi(XD-FMontant.Value,DEV.Decimale) ;
        ResteP:=Arrondi(XP-DeviseToEuro(FMontant.Value,EU.TauxDEV,DEV.Quotite),V_PGI.OkDecV) ;
        //ResteE:=Arrondi(XE-DeviseToPivot(FMontant.Value,EU.TauxDEV,DEV.Quotite),V_PGI.OkDecE) ;
      END else
      BEGIN
        ResteP:=Arrondi(XP-FMontant.Value,V_PGI.OkDecV) ;
        //ResteE:=Arrondi(XE-EuroToPivot(FMontant.Value),V_PGI.OkDecE);
        ResteD:=ResteP ;
      END ;

    Height:=Height+FDateEche2.Top-FTitreSup.Top+30 ;
    FModePaie2.Enabled:=True ;
    FDateEche2.Enabled:=True ;
    FModePaie2.SetFocus ;
    FModePaie.Enabled:=False ;
    FDateEche.Enabled:=False ;
    FMontant.Enabled:=False ;
   END else
   BEGIN
    if Not IsValidDate(FDateEche2.Text) then BEGIN
      HME.Execute(3,'','') ;  // Vous devez renseigner une date valide.
      Exit ;
    END ;
    DDE:=StrToDate(FDateEche2.Text) ;
    if Not NbJoursOK(EU.DateComptable,DDE) then BEGIN
      HME.Execute(4,'','') ;  // La date d'échéance doit respecter la plage de saisie autorisée.
      Exit ;
    END ;
    ModifieEche ;
    if V_PGI.IOError=oeOk then ModalResult:=mrOk ;
  END ;
end;

procedure TFEcheUnit.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((Shift=[]) and (Key=VK_F10)) then BEGIN
    Key:=0 ;
    if BValider.CanFocus then BValider.SetFocus ;
    BValiderClick(Nil) ;
  END ;
end;

procedure TFEcheUnit.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

end.
