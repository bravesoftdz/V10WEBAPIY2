{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit Valperio;

interface

uses Windows,   // PMinMaxInfo
     Messages,  // TMessage, WMGetMinMaxInfo
     Classes,
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     DB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     SysUtils,  // DateToStr, FormatDateTime, StrToDate
     Controls,
     Forms,
     Dialogs,
     Grids,
     Hctrls,
     StdCtrls,
     Buttons,
     ExtCtrls,
     HEnt1,
     Ent1,
     hmsgbox,
     HSysMenu,
     HPanel,
     UiUtil,
     HTB97, Graphics,
     {$IFDEF NETEXPERT}
      UNeActions,
     {$ENDIF}
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     ParamSoc;

Procedure ValidationPeriode ( Validation , Jal : Boolean ) ;
Function ValPerPourClo ( Validation , Jal : Boolean ) : TTraitement ;

Type TInfoExo = Class
        St : String ;
      End ;

type
  TFValperio = class(TForm)
    PExo      : TPanel;
    Pappli    : TPanel;
    TExo      : THLabel;
    CbExo     : THValComboBox;
    Fliste    : THGrid;
    TJal      : THLabel;
    MsgBox    : THMsgBox;
    CBDebDat  : TComboBox;
    CBFinDate : TComboBox;
    CbVal: TCheckBox;
    HMTrad: THSystemMenu;
    BCherche: TToolbarButton97;
    Perjal: THValComboBox;
    Perjal1: THValComboBox;
    PourClotureOk: TCheckBox;
    TPerio: TLabel;
    HPB: TToolWindow97;
    Dock: TDock97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    iCritGlyph: TImage;
    iCritGlyphModified: TImage;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CbExoChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FlisteClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure CbValClick(Sender: TObject);
  private
    FCritModified: Boolean;                                                  {FP 10/11/2005 FQ14941}
    Validation,Jal : Boolean ;
    Exo : TExoDate ;
    NbrMois : Word ;
    PerioJal : String ;
    NbPerio : Integer ;
    WMinX,WMinY : Integer ;
    CtrlExterne : Boolean ;
    procedure SetCritModified(const Value: Boolean);                         {FP 10/11/2005 FQ14941}
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure RempliCbExo ;
    Procedure InitExo ;
    Procedure RempliFliste ;
    Procedure AfficheResultat(St : String ; DebDate : TDateTime) ;
    Procedure RempliLesDates ;
    Function  OnExecute : Boolean ;
    Procedure MajDesTables ;
    Procedure ChargePeriodeJal ;
    Procedure MajPeriodJalVal ;
//    Procedure MajJournaux ;
    Procedure MajPeriodJalDeVal ;

    property CritModified: Boolean read FCritModified write SetCritModified; {FP 10/11/2005 FQ14941}
  public
  end;


implementation

{$R *.DFM}

uses
{$IFDEF EAGLCLIENT}
// A FAIRE Voir BImprimerClick
{$ELSE}
  PrintDBG,
{$ENDIF}
{$IFDEF MODENT1}
  ULibExercice,
{$ENDIF MODENT1}
  Constantes,  {QUALIFTRESO}
  UtilPGI     // _BlocageMonoPoste, _DeblocageMonoPoste
  ;

Procedure ValidationPeriode ( Validation,Jal : Boolean ) ;
var FValperio : TFValperio ;
    PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(True,'',True) then Exit ;
FValperio:=TFValperio.Create(Application) ;
if Validation then FValperio.HelpContext:=7691000
              else FValperio.HelpContext:=7694000 ;
FValperio.Validation:=Validation ; FValperio.CtrlExterne:=FALSE ;
FValperio.Jal:=Jal ;
if FValperio.Jal then FValperio.Width:=346 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FValperio.ShowModal ;
    Finally
     FValperio.Free ;
     _DeblocageMonoPoste(True,'',True) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FValperio,PP) ;
   FValperio.Show ;
   END ;
END ;

Function ValPerPourClo ( Validation , Jal : Boolean ) : TTraitement ;
var FValperio : TFValperio ;
BEGIN
Result:=cPasFait ;
if Not _BlocageMonoPoste(True,'',True) then Exit ;
FValperio:=TFValperio.Create(Application) ;
 Try
  FValperio.Validation:=Validation ;
  FValperio.Jal:=Jal ;
  if FValperio.Jal then FValperio.Width:=346 ;
  FValperio.CtrlExterne:=True ;
  FValperio.ShowModal ;
  Case FValperio.PourClotureOk.State of
    cbGrayed    : Result:=cPasFait ;
    cbChecked   : Result:=cOk ;
    cbUnChecked : Result:=cPasOk ;
    End ;
 Finally
  FValperio.Free ;
  _DeblocageMonoPoste(True,'',True) ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFValperio.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;  
end;

procedure TFValperio.FormShow(Sender: TObject);
begin

{$IFDEF EAGLCLIENT}
    BImprimer.Visible := false;
{$ENDIF}

PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
if Validation then Caption:=MsgBox.Mess[9]+' '+MsgBox.Mess[3]
              else Caption:=MsgBox.Mess[10]+' '+MsgBox.Mess[3] ;
UpdateCaption(Self) ;
RempliCbExo ; ChargePeriodeJal ; PourClotureOk.State:=cbGrayed ;
If CtrlExterne Then CbExo.Value:=VH^.EnCours.Code Else CbExo.Value:=EXRF(VH^.Entree.Code) ;
if Not Validation then BEGIN BValider.Hint:=MsgBox.Mess[5] ; BValider.Caption:=MsgBox.Mess[5] ; END ;
BChercheClick(Nil) ; TPerio.Caption:=MsgBox.Mess[15]+' '+FListe.Cells[0,FListe.Row]
end;

Procedure TFValperio.ChargePeriodeJal ;
Var
  St : String ;
  Q : TQuery;
BEGIN
St:='SELECT J_JOURNAL,J_VALIDEEN,J_VALIDEEN1 FROM JOURNAL WHERE '+
    'J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" AND '+
    'J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA" ORDER BY J_JOURNAL' ;
Q := OpenSQL(St,True);
Perjal.Values.Clear ;
Perjal.Items.Clear ;
Perjal1.Values.Clear ;
Perjal1.Items.Clear ;
  While Not Q.Eof do BEGIN
   Perjal.Values.Add(Q.Fields[0].AsString) ; Perjal1.Values.Add(Q.Fields[0].AsString) ;
   Perjal.Items.Add(Q.Fields[1].AsString) ; Perjal1.Items.Add(Q.Fields[2].AsString) ;
   Q.Next ;
   END ;
Ferme(Q);
END ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/09/2004
Modifié le ... : 23/09/2004
Description .. : - BPY le 23/09/2004 - optimization
Mots clefs ... : 
*****************************************************************}
Procedure TFValperio.MajPeriodJalVal ;
Var
    i,j : Byte;
    St : String;
begin
    if (PerioJal = 'J_VALIDEEN') then
    begin
        for i := 0 to PerJal.Items.Count-1 do
        begin
            // mise a jour de la variable
            St := PerJal.Items[i];
            if (Length(St)<24) then St := St +  StringOfChar('-', 24-Length(St)); // 14691
            for j := 1 to NbPerio do St[j] := 'X';
            PerJal.Items[i] := St;
            // mise a jour de la base
            ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + PerJal.Items[i] + '" Where J_JOURNAL="' + PerJal.Values[i] + '"');
        end;
    end
    else
    begin
        for i := 0 to PerJal1.Items.Count-1 do
        begin
            // mise a jour de la variable
            St := PerJal1.Items[i];
            if (Length(St)<24) then St := St +  StringOfChar('-', 24-Length(St)); // 14691
            for j := 1 to NbPerio do St[j] := 'X';
            PerJal1.Items[i] := St;
            // mise a jour de la base
            ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + PerJal1.Items[i] + '" Where J_JOURNAL="' + PerJal1.Values[i] + '"');
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/09/2004
Modifié le ... : 23/09/2004
Description .. : - BPY le 23/09/2004 - optimization
Mots clefs ... :
*****************************************************************}
Procedure TFValperio.MajPeriodJalDeVal;
var
    i : Byte;
    St : String;
begin
    if (PerioJal = 'J_VALIDEEN') then
    begin
        for i := 0 to PerJal.Items.Count-1 do
        begin
            // mise a jour de la variable
            St := Perjal.Items[i];
            St[NbPerio] := '-';
            Perjal.Items[i] := St;
            // mise a jour de la base
            ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + PerJal.Items[i] + '" Where J_JOURNAL="' + PerJal.Values[i] + '"');
        end;
    end
    else
    begin
        for i := 0 to PerJal1.Items.Count-1 do
        begin
            // mise a jour de la variable
            St := Perjal1.Items[i];
            St[NbPerio] := '-';
            Perjal1.Items[i] := St;
            // mise a jour de la base
            ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + PerJal1.Items[i] + '" Where J_JOURNAL="' + PerJal1.Values[i] + '"');
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/09/2004
Modifié le ... : 23/09/2004
Description .. : - BPY le 23/09/2004 - optimization
Mots clefs ... : 
*****************************************************************}
//Procedure TFValperio.MajJournaux ;
//var
//    i : Byte;
//begin
//    if (PerioJal = 'J_VALIDEEN') then
//    begin
//        for i := 0 to PerJal.Items.Count-1 do
//        begin
//            ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + PerJal.Items[i] + '" Where J_JOURNAL="' + PerJal.Values[i] + '"');
//        end;
//    end
//    else
//    begin
//        for i := 0 to PerJal1.Items.Count-1 do
//        begin
//            ExecuteSQL('UPDATE JOURNAL SET ' + PerioJal + '="' + PerJal1.Items[i] + '" Where J_JOURNAL="' + PerJal1.Values[i] + '"');
//        end;
//    end;
//end;

Procedure TFValperio.RempliCbExo ;
Var Sql : String ;
    X : TInfoExo ;
    Q : TQuery;
BEGIN
  CbExo.Values.Clear ;
  CbExo.Items.Clear ;
//if VH^.Suivant.Code<>'' then Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_DATEDEBUT>="'+USDateTime(VH^.EnCours.Deb)+'"'
if VH^.Suivant.Code<>'' then Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_EXERCICE="'+VH^.EnCours.Code+'" OR EX_EXERCICE="'+VH^.Suivant.Code+'"'
                        else Sql:='Select EX_EXERCICE,EX_LIBELLE,EX_VALIDEE From EXERCICE Where EX_EXERCICE="'+VH^.EnCours.Code+'"' ;
  Q := OpenSQL(Sql,True);
  While Not Q.Eof do BEGIN
    X:=TInfoExo.Create ;
    X.St:=Q.Fields[2].AsString ;
    CbExo.Values.Add(Q.Fields[0].AsString) ;
    CbExo.Items.AddObject(Q.Fields[1].AsString,X) ;
    Q.Next ;
  END ;
  Ferme(Q);
END ;

procedure TFValperio.CbExoChange(Sender: TObject);
begin
if CbExo.Text='' then Exit ;
CritModified := True;         {FP 10/11/2005 FQ14941}
InitExo ;
RempliLesDates ;
end;

Procedure TFValperio.InitExo ;
BEGIN
if CbExo.Value=VH^.Encours.Code then
   BEGIN
   Exo.Code:=VH^.Encours.Code ; Exo.Deb:=VH^.Encours.Deb ;
   Exo.Fin:=VH^.Encours.Fin ; NbrMois:=VH^.Encours.NombrePeriode ; PerioJal:='J_VALIDEEN' ;
   END else
   BEGIN
   Exo.Code:=VH^.Suivant.Code ; Exo.Deb:=VH^.Suivant.Deb ;
   Exo.Fin:=VH^.Suivant.Fin ; NbrMois:=VH^.Suivant.NombrePeriode ; PerioJal:='J_VALIDEEN1' ;
   END ;
END ;

Procedure TFValperio.RempliFliste ;
BEGIN
FListe.VidePile(True) ;
if TInfoExo(CbExo.Items.Objects[CbExo.ItemIndex]).St='' then Exit ;
AfficheResultat(TInfoExo(CbExo.Items.Objects[CbExo.ItemIndex]).St,Exo.Deb) ; FlisteClick(Nil) ;
if FListe.RowCount>2 then Fliste.RowCount:=Fliste.RowCount-1 ;
END;

Procedure TFValperio.AfficheResultat(St : String ; DebDate : TDateTime) ;
Var i,Ind : Byte ;
    Okok  : boolean ;
    StDate : String ;
    X : TInfoExo ;
BEGIN
for i:=1 to NbrMois do
    BEGIN
    Okok:=False ; Ind:=0 ;
    if CbVal.State=cbGrayed then BEGIN Okok:=True ; if St[i]='X' then Ind:=0 else Ind:=1 ; END else
     if (CbVal.State=cbChecked) And (St[i]='X') then BEGIN Okok:=True ; Ind:=0 ; END else
      if (CbVal.State=cbUnchecked) And (St[i]='-')  then BEGIN Okok:=True ; Ind:=1 ; END ;
    if Okok then
       BEGIN
       X:=TInfoExo.Create ; X.St:=DateToStr(DebDate) ;
       StDate:=FormatDateTime('mmmm-yyyy',DebDate) ; FirstMajuscule(StDate) ;
       FListe.Cells[0,Fliste.RowCount-1]:=StDate ; FListe.Objects[0,Fliste.RowCount-1]:=X ;
       FListe.Cells[1,Fliste.RowCount-1]:=MsgBox.Mess[Ind] ;
       Fliste.RowCount:=Fliste.RowCount+1 ; ;
       END ;
    DebDate:=PlusMois(DebDate,1) ;
    END ;
END ;

Procedure TFValperio.RempliLesDates ;
Var i : Byte ;
    Dda,Fda : TDateTime ;
BEGIN
CBDebDat.Items.Clear ; CBFinDate.Items.Clear ;
Dda:=Exo.Deb ; Fda:=FindeMois(Dda) ;
for i:=1 to NbrMois do
    BEGIN
    CBDebDat.Items.Add(DatetoStr(Dda)) ; CBFinDate.Items.Add(DatetoStr(Fda)) ;
    Dda:=PlusMois(Dda,1) ; Fda:=FindeMois(Dda) ;
    END ;
END ;

procedure TFValperio.BValiderClick(Sender: TObject);
begin
  NbPerio:=CBDebDat.Items.IndexOf(TInfoExo(FListe.Objects[0,FListe.Row]).St)+1 ;
  if Not OnExecute then
  begin
    MsgBox.Execute(6,caption,'') ;
    Exit ;
  end ;

  if Validation then
  begin
    If MsgBox.Execute(7,caption,'')=mrYes then MajDesTables ;
  end
  else
  begin
{$IFNDEF CERTIFNF}
    if GetParamSocSecur('SO_CPCONFORMEBOI', False) then //06/12/2006 YMO Norme NF 203
    begin
      PgiInfo('Pour la conformité stricte avec la norme NF 203 (et le BOI du 24/01/2006) cette fonction n''est plus disponible',Caption);
      Exit;
    end;
{$ENDIF}

    if MsgBox.Execute(8,caption,'')=mrYes then
    begin
      PgiInfo('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 23 qui rappelle l''interdiction de ' + #13#10 +
            'toute modification et/ou suppression d''écritures validées, nous vous conseillons de « tracer » l’information ' + #13#10 +
            'par toute méthode à votre disposition (par exemple le bloc notes).', Caption);
      MajDesTables ;
    end;
  end;
  BChercheClick(Nil) ;
end;

Function TFValperio.OnExecute : Boolean ;
BEGIN
Result:=False ;
if Validation then
   BEGIN
   if FListe.Cells[1,FListe.Row]=MsgBox.Mess[1] then Result:=True ;
   END else
   BEGIN
   if FListe.Cells[1,FListe.Row]=MsgBox.Mess[0] then Result:=True ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/09/2004
Modifié le ... : 23/09/2004
Description .. : - BPY 23/09/2004 - optimisation suite a la fiche de bug n°
Suite ........ : 14635
Mots clefs ... : 
*****************************************************************}
Procedure TFValperio.MajDesTables ;
Var
    Sql,St : String;
    i : Byte;
    DateD,DateF : String;
    ToutesValides : Boolean;
    Unexo : TExodate;
    ListePiece : HTStringList;
    QPiece : TQuery;
    {$IFDEF NETEXPERT}
    IsNetEXpert : Boolean;
    {$ENDIF}
begin
    {$IFDEF NETEXPERT}
    if not InterrogeUserService (IsNetEXpert) then
    begin
        if Validation and IsNetExpert then
        begin
            PGIInfo('Il y a des utilisateurs connectés sur le dossier Business Line, cette fonction n''est plus disponible',Caption);
            Exit;
        end;
    end;
    {$ENDIF}

    ListePiece:=HTStringList.Create ;
    ListePiece.Clear ;
    try
    {$IFDEF CERTIFNF}
      If Validation Then
         ListePiece.Add('VALECR ')
      Else
         ListePiece.Add('ANNVALECR ');
    {$ELSE}
      ListePiece.Sorted:=True ;
    {$ENDIF}

      if (Validation) then
      begin
          // Init des dates
          DateD := UsDateTime(StrToDate(CBDebDat.Items[0]));
          DateF := UsDateTime(StrToDate(CBFinDate.Items[NbPerio-1]));

          // mise a jour de la variable de validation
          St := TInfoExo(CbExo.Items.Objects[CbExo.ItemIndex]).St;
          for i := 1 to NbPerio do St[i] := 'X';
          TInfoExo(CbExo.Items.Objects[CbExo.ItemIndex]).St := St;

          // mise a jour de l'exercice
          ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + St + '" WHERE EX_EXERCICE="' + CbExo.Value + '"');

          // mise a jour des ecritures
          Sql := 'UPDATE ECRITURE SET E_VALIDE="X" WHERE E_EXERCICE="' + CbExo.Value + '" AND E_DATECOMPTABLE>="' + DateD
               + '" AND E_DATECOMPTABLE<="' + DateF + '" AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I") AND E_ECRANOUVEAU="N" '; // SBO 19/08/2004 FQ14209 prise en compte des écritures "I"
          ExecuteSQL(Sql);

          QPiece:=OpenSQL(StringReplace(Sql,'UPDATE ECRITURE SET E_VALIDE="X"','SELECT E_JOURNAL,E_PERIODE,E_NUMEROPIECE FROM ECRITURE',[]),True) ;
          While Not QPiece.EOF do
          BEGIN
             ListePiece.Add(QPiece.Fields[0].AsString+'-'+QPiece.Fields[1].AsString+'-'+QPiece.Fields[2].AsString);
             QPiece.Next ;
          END ;
          Ferme(QPiece) ;

          // mise a jour des ventilations // FQ16762 SBO 04/10/2005 Prise en compte des ventilations à la validation
          Sql := 'UPDATE ANALYTIQ SET Y_VALIDE="X" WHERE Y_EXERCICE="' + CbExo.Value + '" AND Y_DATECOMPTABLE>="' + DateD
               + '" AND Y_DATECOMPTABLE<="' + DateF + '" AND (Y_QUALIFPIECE="N" OR Y_QUALIFPIECE="I") '
               + 'AND Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N" '; // SBO 19/08/2004 FQ14209 prise en compte des écritures "I"
          ExecuteSQL(Sql);

          if St <> '' then
          begin
              // recup de l'exercice
              UnExo.Code := CbExo.Value;
              RempliExoDate(UnExo);

              // routine pour savoir si toutes les periodes de l'exo sont validees
              ToutesValides := True;
              for i := 1 to UnExo.NombrePeriode do
                  if (St[i] = '-') then
                  begin
                      ToutesValides := false;
                      Break;
                  end;

              // si elle le sont ....
              if (ToutesValides) then PourClotureOk.State := cbChecked
              else PourClotureOk.State := cbUnChecked;
          end;

          // GCO - 18/09/2006
          // YMO 07/12/2006 Norme NF Suppression
          // CPEnregistreLog('VALPERIODE ' + FormatDateTime('mmm yyyy', StrToDate(CBFinDate.Items[NbPerio-1])));

          // mise a jour des journaux
          MajPeriodJalVal;
    {$IFDEF NETEXPERT}
             // Synchronisation avec BL
          if IsNetExpert then
                SynchroOutBL ('V', CBFinDate.Items[NbPerio-1], '');
    {$ENDIF}

    //        MajJournaux;
      end
      else
      begin
          // test : on ne devalide pas une periode cloturé !
          if (StrToDate(CBDebDat.Items[NbPerio-1]) <= VH^.DateCloturePer) then
          begin
              MsgBox.Execute(16,Caption,'');
              Exit;
          end;

          // Init des dates
          //DateD := UsDateTime(StrToDate(CBDebDat.Items[NbPerio-1]));
          { FQ 16139 - CA - 11/07/2005 : Si date début d'exercice <> 1er jour du mois, comparer par rapport
          au premier jour du mois à partir du deuxième mois de l'exercice }
          if (NbPerio > 1) then DateD := UsDateTime(DebutDeMois(StrToDate(CBDebDat.Items[NbPerio-1])))
          else  DateD := UsDateTime(StrToDate(CBDebDat.Items[NbPerio-1]));
          DateF := UsDateTime(StrToDate(CBFinDate.Items[NbPerio-1]));

          // mise a jour de la variable de validation
          TInfoExo(CbExo.Items.Objects[CbExo.ItemIndex]).St[NbPerio] := '-';
          St := TInfoExo(CbExo.Items.Objects[CbExo.ItemIndex]).St;

          // mise a jour de l'exercice
          ExecuteSQL('UPDATE EXERCICE SET EX_VALIDEE="' + St + '" WHERE EX_EXERCICE="' + CbExo.Value + '"');

          // mise a jour des ecritures
          Sql := 'UPDATE ECRITURE SET E_VALIDE="-" WHERE E_EXERCICE="' + CbExo.Value + '" AND E_DATECOMPTABLE>="' + DateD
               + '" AND E_DATECOMPTABLE<="' + DateF + '" AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="I") AND E_ECRANOUVEAU="N" AND E_CREERPAR<>"DET"' ; // SBO 19/08/2004 FQ14209 prise en compte des écritures "I"
          { CA - 10/10/2003 - FQ 12448 - On ne touche pas aux pièces validées par la gescom
            JP - 29/07/05 : FQ 15124 : On revient en arrière
          Sql := Sql + ' AND E_REFGESCOM=""';
           JP 26/06/07 : Par contre on ne touche pas aux pièces de Tréso}
          Sql := Sql + ' AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL OR E_QUALIFORIGINE = "")';
          ExecuteSQL(Sql);

          QPiece:=OpenSQL(StringReplace(Sql,'UPDATE ECRITURE SET E_VALIDE="-"','SELECT E_JOURNAL,E_PERIODE,E_NUMEROPIECE FROM ECRITURE',[]),True) ;
          While Not QPiece.EOF do
          BEGIN
             ListePiece.Add(QPiece.Fields[0].AsString+'-'+QPiece.Fields[1].AsString+'-'+QPiece.Fields[2].AsString);
             QPiece.Next ;
          END ;
          Ferme(QPiece) ;

          // mise a jour des ventilations // FQ16762 SBO 04/10/2005 Prise en compte des ventilations à la validation
          Sql := 'UPDATE ANALYTIQ SET Y_VALIDE="-" WHERE Y_EXERCICE="' + CbExo.Value + '" AND Y_DATECOMPTABLE>="' + DateD
               + '" AND Y_DATECOMPTABLE<="' + DateF + '" AND (Y_QUALIFPIECE="N" OR Y_QUALIFPIECE="I") '
               + 'AND Y_TYPEANALYTIQUE="-" AND Y_ECRANOUVEAU="N" ';
          ExecuteSQL(Sql);

          // GCO - 18/09/2006
          // YMO 07/12/2006 Norme NF Suppression
          // CPEnregistreLog('ANNULVALPERIODE ' + FormatDateTime('mmm yyyy', StrToDate(CBFinDate.Items[NbPerio-1])));

          // mise a jour des journaux
          MajPeriodJalDeVal;
    //        MajJournaux;
      end;
    except
      on E: Exception do
      begin
{$IFNDEF EAGLSERVER}
         PgiError(E.Message);
{$ENDIF}
         if assigned(QPiece) then Ferme(QPiece);
      end;
    end;

    { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
    For i:=0 To ListePiece.Count-1 Do
    Begin
     If Validation Then
        CPEnregistreLog('VALECR ' + ListePiece[i])
     Else
        CPEnregistreLog('ANNVALECR ' + ListePiece[i]);
    End;
{$ELSE}
    CPEnregistreJalEvent('CVE','Validation d''écritures',ListePiece);
{$ENDIF}

    ListePiece.Free ;

    // recherche ....
    BChercheClick(Nil);
end;

procedure TFValperio.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  // A FAIRE
{$ELSE}
  PrintDBGrid(FListe,PExo,Caption,'') ;
{$ENDIF}
end;

procedure TFValperio.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ;
if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True,'',True) ;
{$IFDEF EAGLCLIENT}
{$ELSE}
   Action:=caFree ;
{$ENDIF}
   END ;
end;

procedure TFValperio.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFValperio.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
CritModified := false;         {FP 10/11/2005 FQ14941}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 23/09/2004
Modifié le ... : 23/09/2004
Description .. : - BPY le 23/09/2004 - optimization
Mots clefs ... : 
*****************************************************************}
procedure TFValperio.BChercheClick(Sender: TObject);
begin
    CritModified := false;         {FP 10/11/2005 FQ14941}
    // nettoyage de la grille des journaux
    FListe.ClearSelected;

    if (CbExo.Text <> '') then RempliFliste;
end;

procedure TFValperio.FlisteClick(Sender: TObject);
begin
TPerio.Caption:=MsgBox.Mess[15]+' '+FListe.Cells[0,FListe.Row]
end;

procedure TFValperio.FormResize(Sender: TObject);
begin
FListe.ColWidths[1]:=FListe.Width -155 ;
end;

procedure TFValperio.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFValperio.SetCritModified(const Value: Boolean);
begin
  {b FP 10/11/2005 FQ14941}
  if FCritModified <> Value then
    begin
    if Value then
      BCherche.Glyph := iCritGlyphModified.Picture.BitMap
    else
      BCherche.Glyph := iCritGlyph.Picture.BitMap;
    end;
  FCritModified := Value;
  {e FP 10/11/2005 FQ14941}
end;

procedure TFValperio.CbValClick(Sender: TObject);
begin
CritModified := True;         {FP 10/11/2005 FQ14941}
end;

end.
