unit AssistMajTarif;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ComCtrls, ExtCtrls,
  UTOB, HPanel, Mask, HEnt1,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  TarifUtil, UtilGC;

  procedure Assist_MajTarif (TOBT : TOB);


type
  TFAssistTarif = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    GBMAJ: TGroupBox;
    CB_PRIX: TCheckBox;
    CB_DATE: TCheckBox;
    TCGPRIX: THLabel;
    TCBDATE: THLabel;
    TINTRO: THLabel;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    PBevel2: TBevel;
    PBevel1: TBevel;
    PBevel3: TBevel;
    GBPRIX: TGroupBox;
    TTYPEMAJPRIX: THLabel;
    TYPEMAJPRIX: THValComboBox;
    VALEURPRIX: THNumEdit;
    TVALEURPRIX: THLabel;
    GBREMISE: TGroupBox;
    TYPEMAJREMISE: THValComboBox;
    TTYPEMAJREMISE: THLabel;
    VALEURREMISE: THNumEdit;
    TVALEURREMISE: THLabel;
    TARRONDIPRIX: THLabel;
    ARRONDIPRIX: THValComboBox;
    DATEEFFET: THCritMaskEdit;
    TDATEEFFET: THLabel;
    GBPERIODE: TGroupBox;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    PBevel4: TBevel;
    TRecap: THLabel;
    ListRecap: TListBox;
    TYPEMAJPERIODE: THValComboBox;
    TTYPEMAJPERIODE: THLabel;
    NBREJOUR: THNumEdit;
    UpDownJour: TUpDown;
    TNBREJOUR: THLabel;
    CBDATEDEBUT: TCheckBox;
    CBDATEFIN: TCheckBox;
    DATEFIN: THCritMaskEdit;
    DATEDEBUT: THCritMaskEdit;
    TDATEDEBUT: THLabel;
    TDATEFIN: THLabel;
    HRecap: THMsgBox;
    HMsgErr: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure CB_PRIXClick(Sender: TObject);
    procedure CB_DATEClick(Sender: TObject);
    procedure TYPEMAJPRIXChange(Sender: TObject);
    procedure TYPEMAJREMISEChange(Sender: TObject);
    procedure TYPEMAJPERIODEChange(Sender: TObject);
    procedure CBDATEDEBUTClick(Sender: TObject);
    procedure CBDATEFINClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure DATEEFFETExit(Sender: TObject);
    procedure DATEDEBUTExit(Sender: TObject);
    procedure DATEFINExit(Sender: TObject);
  private
    { Déclarations privées }
    TOBTarif : TOB ;
//    function  SpaceStr ( nb : integer) : string;
//    function  ExtractLibelle ( St : string) : string;
    procedure ListeRecap;
// Traitement
    procedure TraiterTarif ;
    procedure MajTarif ;
    procedure MajTarifDetail  (TOBD : TOB);
    procedure ModifPrix (Var TOBD : TOB; ModifDateEffet : Boolean) ;
    procedure ModifRemise (Var TOBD : TOB; ModifDateEffet : Boolean) ;
    procedure ModifPeriode (Var TOBD : TOB) ;
  public
    { Déclarations publiques }
  end;

var
  FAssistTarif: TFAssistTarif;
  i_NumEcran : integer;

implementation

{$R *.DFM}

procedure Assist_MajTarif (TOBT : TOB);
var
   Fo_Assist : TFAssistTarif;
Begin
     Fo_Assist := TFAssistTarif.Create (Application);
     Fo_Assist.TOBTarif:= TOBT;
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

{=========================================================================================}
{============================= Evenements de la forme ====================================}
{=========================================================================================}
procedure TFAssistTarif.FormShow(Sender: TObject);
var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
bAnnuler.Visible := True;
bSuivant.Enabled := False;
bFin.Visible := True;
bFin.Enabled := False;
//i_NumEcran := 0;

Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;

end;

{=========================================================================================}
{================================= Récapitulatif =========================================}
{=========================================================================================}
(* function TFAssistTarif.SpaceStr ( nb : integer) : string;
Var St_Chaine : string ;
    i_ind : integer ;
BEGIN
St_Chaine := '' ;
for i_ind := 1 to nb do St_Chaine:=St_chaine+' ';
Result:=St_Chaine;
END;

function TFAssistTarif.ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
BEGIN
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    BEGIN
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    END else St_Chaine := St;
Result := St_Chaine + ' : ';
END; *)

procedure TFAssistTarif.ListeRecap;
Var st_chaine : string;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add (PTITRE.Caption);
ListRecap.Items.Add ('');
if CB_PRIX.checked then
    BEGIN
    st_chaine := HRecap.Mess[0];
    ListRecap.Items.Add (ExtractLibelle (CB_PRIX.Caption) + st_chaine);
    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TDATEEFFET.Caption) + DATEEFFET.Text);
    ListRecap.Items.Add ( SpaceStr(4) + ExtractLibelle (GBPRIX.Caption));
    ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TTYPEMAJPRIX.Caption) + TYPEMAJPRIX.Text);
    if TYPEMAJPRIX.Value <> '' then
        BEGIN
        ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TVALEURPRIX.Caption) + VALEURPRIX.Text);
        ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TARRONDIPRIX.Caption) + ARRONDIPRIX.Text);
        END;
    ListRecap.Items.Add ( SpaceStr(4) + ExtractLibelle (GBREMISE.Caption));
    ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TTYPEMAJREMISE.Caption) + TYPEMAJREMISE.Text);
    if TYPEMAJREMISE.Value <> '' then
        BEGIN
        ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TVALEURREMISE.Caption) + VALEURREMISE.Text);
        END;
    END else
    BEGIN
    st_chaine := HRecap.Mess[1];
    ListRecap.Items.Add (ExtractLibelle (CB_PRIX.Caption) + st_chaine);
    END;
ListRecap.Items.Add ('');
if CB_DATE.checked then
    BEGIN
    st_chaine := HRecap.Mess[0];
    ListRecap.Items.Add (ExtractLibelle (CB_DATE.Caption) + st_chaine);
    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TTYPEMAJPERIODE.Caption) + TYPEMAJPERIODE.Text);
    if TYPEMAJPERIODE.Value <> '' then
        if TYPEMAJPERIODE.Value = 'D01' then
            BEGIN
            if CBDATEDEBUT.Checked then
                BEGIN
                st_chaine := HRecap.Mess[0];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEDEBUT.Caption) + St_Chaine);
                ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TDATEDEBUT.Caption) + DATEDEBUT.Text);
                END else
                BEGIN
                st_chaine := HRecap.Mess[1];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEDEBUT.Caption) + St_Chaine);
                END;
            if CBDATEFIN.Checked then
                BEGIN
                st_chaine := HRecap.Mess[0];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEFIN.Caption) + St_Chaine);
                ListRecap.Items.Add (SpaceStr(8) + ExtractLibelle (TDATEFIN.Caption) + DATEFIN.Text);
                END else
                BEGIN
                st_chaine := HRecap.Mess[1];
                ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (CBDATEFIN.Caption) + St_Chaine);
                END;
            END else
            BEGIN
            ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TNBREJOUR.Caption) + NBREJOUR.Text);
            END;
    END else
    BEGIN
    st_chaine := HRecap.Mess[1];
    ListRecap.Items.Add (ExtractLibelle (CB_DATE.Caption) + st_chaine);
    END;
ListRecap.Items.Add ('');
END;

{=========================================================================================}
{========================== Evenements Liés aux Boutons ==================================}
{=========================================================================================}
procedure TFAssistTarif.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if i_NumEcran <= 2 then NBREJOUR.Text:=FloatToStr(NBREJOUR.Value); 
if (CB_PRIX.Checked = False) and (i_NumEcran = 1) then
    BEGIN
    RestorePage ;
    Onglet := NextPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    NBREJOUR.Text:=FloatToStr(NBREJOUR.Value);
    END;
if (CB_DATE.Checked = False) and (i_NumEcran = 2) then
    BEGIN
    RestorePage ;
    Onglet := NextPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    END;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
if bFin.Enabled then ListeRecap;

end;

procedure TFAssistTarif.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if i_NumEcran = 2 then NBREJOUR.Text:=FloatToStr(NBREJOUR.Value); 
if (CB_PRIX.Checked = False) and (i_NumEcran = 1) then
    BEGIN
    RestorePage ;
    Onglet := PreviousPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    END;
if (CB_DATE.Checked = False) and (i_NumEcran = 2) then
    BEGIN
    RestorePage ;
    Onglet := PreviousPage ;
    if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
    END;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;

end;

procedure TFAssistTarif.bFinClick(Sender: TObject);
begin
  inherited;
TraiterTarif ;
Close ;
end;

{=========================================================================================}
{============================= Evenements Onglet 1 =======================================}
{=========================================================================================}
procedure TFAssistTarif.CB_PRIXClick(Sender: TObject);
begin
  inherited;
bSuivant.Enabled:=(CB_PRIX.Checked) or (CB_DATE.Checked);
end;

procedure TFAssistTarif.CB_DATEClick(Sender: TObject);
begin
  inherited;
bSuivant.Enabled:=(CB_PRIX.Checked) or (CB_DATE.Checked);
end;

{=========================================================================================}
{============================= Evenements Onglet 2 =======================================}
{=========================================================================================}
procedure TFAssistTarif.TYPEMAJPRIXChange(Sender: TObject);
begin
  inherited;
if TYPEMAJPRIX.Value <> '' then
    BEGIN
    VALEURPRIX.Enabled := True;
    VALEURPRIX.Color := clWindow;
    VALEURPRIX.TabStop := True;
    if TYPEMAJPRIX.Value = 'P03' then VALEURPRIX.NumericType:= ntPercentage
                                 else VALEURPRIX.NumericType:= ntGeneral;
    END else
    BEGIN
    VALEURPRIX.Enabled := False;
    VALEURPRIX.Color := clBtnFace;
    VALEURPRIX.TabStop := False;
    END;
end;

procedure TFAssistTarif.TYPEMAJREMISEChange(Sender: TObject);
begin
  inherited;
if TYPEMAJREMISE.Value <> '' then
    BEGIN
    VALEURREMISE.Enabled := True;
    VALEURREMISE.Color := clWindow;
    VALEURREMISE.TabStop := True;
    END else
    BEGIN
    VALEURREMISE.Enabled := False;
    VALEURREMISE.Color := clBtnFace;
    VALEURREMISE.TabStop := False;
    END;
end;

procedure TFAssistTarif.DATEEFFETExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DATEEFFET.Text) then
  begin
    stErr:='"'+ DATEEFFET.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DATEEFFET.SetFocus;
  end;
end;

{=========================================================================================}
{============================= Evenements Onglet 3 =======================================}
{=========================================================================================}

procedure TFAssistTarif.TYPEMAJPERIODEChange(Sender: TObject);
begin
  inherited;
if (TYPEMAJPERIODE.Value = 'D02') or (TYPEMAJPERIODE.Value = 'D03') then
    BEGIN
    NBREJOUR.Enabled := True;
    NBREJOUR.Color := clWindow;
    NBREJOUR.TabStop := True;
    UpDownJour.Enabled := True;
    CBDATEDEBUT.Enabled := False;
    CBDATEDEBUT.TabStop := False;
    CBDATEFIN.Enabled := False;
    CBDATEFIN.TabStop := False;
    DATEDEBUT.Enabled := False;
    DATEDEBUT.Color := clBtnFace;
    DATEDEBUT.TabStop := False;
    DATEFIN.Enabled := False;
    DATEFIN.Color := clBtnFace;
    DATEFIN.TabStop := False;
    END else if TYPEMAJPERIODE.Value = 'D01' then
    BEGIN
    NBREJOUR.Enabled := False;
    NBREJOUR.Color := clBtnFace;
    NBREJOUR.TabStop := False;
    UpDownJour.Enabled := False;
    CBDATEDEBUT.Enabled := True;
    CBDATEDEBUT.TabStop := True;
    CBDATEFIN.Enabled := True;
    CBDATEFIN.TabStop := True;
    if CBDATEDEBUT.Checked then
        BEGIN
        DATEDEBUT.Enabled := True;
        DATEDEBUT.Color := clWindow;
        DATEDEBUT.TabStop := True;
        END else
        BEGIN
        DATEDEBUT.Enabled := False;
        DATEDEBUT.Color := clBtnFace;
        DATEDEBUT.TabStop := False;
        END;
    if CBDATEFIN.Checked then
        BEGIN
        DATEFIN.Enabled := True;
        DATEFIN.Color := clWindow;
        DATEFIN.TabStop := True;
        END else
        BEGIN
        DATEFIN.Enabled := False;
        DATEFIN.Color := clBtnFace;
        DATEFIN.TabStop := False;
        END;
    END else
    BEGIN
    NBREJOUR.Enabled := False;
    NBREJOUR.Color := clBtnFace;
    NBREJOUR.TabStop := False;
    UpDownJour.Enabled := False;
    CBDATEDEBUT.Enabled := False;
    CBDATEDEBUT.TabStop := False;
    CBDATEFIN.Enabled := False;
    CBDATEFIN.TabStop := False;
    DATEDEBUT.Enabled := False;
    DATEDEBUT.Color := clBtnFace;
    DATEDEBUT.TabStop := False;
    DATEFIN.Enabled := False;
    DATEFIN.Color := clBtnFace;
    DATEFIN.TabStop := False;
    END;
end;

procedure TFAssistTarif.CBDATEDEBUTClick(Sender: TObject);
begin
  inherited;
if CBDATEDEBUT.Checked then
    BEGIN
    DATEDEBUT.Enabled := True;
    DATEDEBUT.Color := clWindow;
    DATEDEBUT.TabStop := True;
    END else
    BEGIN
    DATEDEBUT.Enabled := False;
    DATEDEBUT.Color := clBtnFace;
    DATEDEBUT.TabStop := False;
    END;
end;

procedure TFAssistTarif.CBDATEFINClick(Sender: TObject);
begin
  inherited;
if CBDATEFIN.Checked then
    BEGIN
    DATEFIN.Enabled := True;
    DATEFIN.Color := clWindow;
    DATEFIN.TabStop := True;
    END else
    BEGIN
    DATEFIN.Enabled := False;
    DATEFIN.Color := clBtnFace;
    DATEFIN.TabStop := False;
    END;
end;

procedure TFAssistTarif.DATEDEBUTExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DATEDEBUT.Text) and (DATEDEBUT.Enabled) then
  begin
    stErr:='"'+ DATEDEBUT.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DATEDEBUT.SetFocus;
  end;
end;

procedure TFAssistTarif.DATEFINExit(Sender: TObject);
var stErr : string;
begin
  inherited;
  if not IsValidDate(DATEFIN.Text) and (DATEFIN.Enabled) then
  begin
    stErr:='"'+ DATEFIN.Text + '" ' + HMsgErr.Mess[0];
    HShowMessage('0;'+PTITRE.Caption+';'+StErr+';W;O;O;O;','','') ;
    DATEFIN.SetFocus;
  end;
end;

{=========================================================================================}
{================================== Traitement ===========================================}
{=========================================================================================}

procedure TFAssistTarif.TraiterTarif ;
Var ioerr : TIOErr ;
    TOBJnal : TOB ;
    QQ : TQuery ;
    NumEvt,NbEnreg : integer ;
BEGIN
NbEnreg:=TOBTarif.detail.count ;
if NbEnreg <= 0 then exit;

NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'TAR');
TOBJnal.PutValue('GEV_LIBELLE', PTITRE.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
ioerr := Transactions (MajTarif, 2);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
Case ioerr of
        oeOk  : BEGIN
                TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
                ListRecap.Items.Add('');
                ListRecap.Items.Add(IntToStr (NbEnreg) + ' ' + HRecap.Mess[4]);
                TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
                END;
    oeUnknown : BEGIN
                MessageAlerte(HRecap.Mess[2]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(HRecap.Mess[2]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
    oeSaisie  : BEGIN
                MessageAlerte(HRecap.Mess[3]) ;
                ListRecap.Items.Add('');
                ListRecap.Items.Add(HRecap.Mess[3]);
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
   END ;
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
END;

procedure TFAssistTarif.MajTarif ;
Var i_ind : integer ;
    TOBD : TOB ;
BEGIN
for i_ind := 0 to TOBTarif.detail.count -1 do
    BEGIN
    TOBD := TOBTarif.Detail[i_ind] ;
    MajTarifDetail (TOBD) ;
    END;
TOBTarif.UpdateDB(False) ;
END;

procedure TFAssistTarif.MajTarifDetail (TOBD : TOB) ;
Var ModifDateEffet : Boolean ;
BEGIN
ModifDateEffet:= True;
if CB_PRIX.checked then
    BEGIN
    if TYPEMAJPRIX.Value <> '' then
        BEGIN
        ModifPrix (TOBD, ModifDateEffet);
        ModifDateEffet:= False;
        END;
    if TYPEMAJREMISE.Value <> '' then
        BEGIN
        ModifRemise (TOBD, ModifDateEffet);
        //ModifDateEffet:= False;
        END;
    END;
if CB_DATE.checked then
    BEGIN
    if TYPEMAJPERIODE.Value <> '' then
        BEGIN
        ModifPeriode (TOBD);
        END;
    END;
END;

procedure TFAssistTarif.ModifPrix (Var TOBD : TOB; ModifDateEffet : Boolean) ;
Var Valeur, Prix : Double;
    CodeArr, CodeDevise : String ;
BEGIN
if TOBD.GetValue ('GF_REGIMEPRIX')='GLO' then exit;
Valeur:= VALEURPRIX.Value;
CodeArr := ARRONDIPRIX.Value;
CodeDevise := TOBD.GetValue ('GF_DEVISE');
if ModifDateEffet then
    BEGIN
    TOBD.PutValue ('GF_PRIXANCIEN', TOBD.GetValue ('GF_PRIXUNITAIRE'));
    TOBD.PutValue ('GF_REMISEANCIEN', TOBD.GetValue ('GF_REMISE'));
    TOBD.PutValue ('GF_DATEEFFET', StrToDate (DATEEFFET.Text));
    END;
if TYPEMAJPRIX.Value = 'P01' then
    BEGIN
    TOBD.PutValue ('GF_PRIXUNITAIRE', Valeur) ;
    END else if TYPEMAJPRIX.Value = 'P02' then
        BEGIN
        Prix := TOBD.GetValue ('GF_PRIXUNITAIRE') * Valeur;
        Prix := ArrondirPrix (CodeArr, Prix);
        TOBD.PutValue ('GF_PRIXUNITAIRE', Prix) ;
        END else if TYPEMAJPRIX.Value = 'P03' then
            BEGIN
            Prix := (TOBD.GetValue ('GF_PRIXUNITAIRE') * (1 + Valeur));
            Prix := ArrondirPrix (CodeArr, Prix);
            TOBD.PutValue ('GF_PRIXUNITAIRE', Prix) ;
            END;
END;

procedure TFAssistTarif.ModifRemise (Var TOBD : TOB; ModifDateEffet : Boolean) ;
Var Valeur, Remise : Double;
    Calcremise : string ;
BEGIN
Valeur:= strToFloat(VALEURREMISE.Text);
if ModifDateEffet then
    BEGIN
    TOBD.PutValue ('GF_PRIXANCIEN', TOBD.GetValue ('GF_PRIXUNITAIRE'));
    TOBD.PutValue ('GF_REMISEANCIEN', TOBD.GetValue ('GF_REMISE'));
    TOBD.PutValue ('GF_DATEEFFET', StrToDate (DATEEFFET.Text));
    END;
if TYPEMAJREMISE.Value = 'R01' then
    BEGIN
    TOBD.PutValue ('GF_REMISE', Valeur) ;
    TOBD.PutValue ('GF_CALCULREMISE', FloatToStr (Valeur)) ;
    END else if TYPEMAJREMISE.Value = 'R02' then
        BEGIN
        CalcRemise := TOBD.GetValue ('GF_CALCULREMISE') + '+' + Trim(FloatToStr (Valeur));
        Remise := RemiseResultante (CalcRemise);
        TOBD.PutValue ('GF_REMISE', Remise) ;
        TOBD.PutValue ('GF_CALCULREMISE', CalcRemise) ;
        END
END;

procedure TFAssistTarif.ModifPeriode (Var TOBD : TOB) ;
Var DateD, DateF : TDateTime ;
    nbj : integer ;
BEGIN
if TYPEMAJPERIODE.Value = 'D01' then
    BEGIN
    if CBDATEDEBUT.Checked then DateD:=StrToDate(DATEDEBUT.Text) else DateD:=TOBD.GetValue('GF_DATEDEBUT') ;
    if CBDATEFIN.Checked then Datef:=StrToDate(DATEFIN.Text) else DateF:=TOBD.GetValue('GF_DATEFIN') ;
    if DateD <= DateF then
        BEGIN
        TOBD.PutValue ('GF_DATEDEBUT', DateD);
        TOBD.PutValue ('GF_DATEFIN', DateF);
        END;
    END else if TYPEMAJPERIODE.Value = 'D02' then
        BEGIN
        DateD:=TOBD.GetValue('GF_DATEDEBUT') ;
        DateF:=TOBD.GetValue('GF_DATEFIN') ;
        nbj := StrToInt (NBREJOUR.Text) ;
        DateF:=DateF + nbj;
        if DateF > IDate2099 then  DateF := IDate2099 ;
        if DateD <= DateF then
            BEGIN
            TOBD.PutValue ('GF_DATEDEBUT', DateD);
            TOBD.PutValue ('GF_DATEFIN', DateF);
            END;
        END else if TYPEMAJPERIODE.Value = 'D03' then
        BEGIN
        DateD:=TOBD.GetValue('GF_DATEDEBUT') ;
        DateF:=TOBD.GetValue('GF_DATEFIN') ;
        nbj := StrToInt (NBREJOUR.Text) ;
        DateF:=DateF - nbj;
        if DateD <= DateF then
            BEGIN
            TOBD.PutValue ('GF_DATEDEBUT', DateD);
            TOBD.PutValue ('GF_DATEFIN', DateF);
            END;
        END;
END;

end.
