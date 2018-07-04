{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/12/2000
Modifié le ... :  20/12/2000
Description .. : Source TOF de la TABLE : SAISECART ()
Mots clefs ... : TOF;SAISECART
*****************************************************************}
Unit TOFSAISECART ;

Interface

Uses Windows, StdCtrls, Controls, Classes, db, forms, sysutils, ComCtrls,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF, ent1, Graphics, utob, SAISUTIL, paramsoc,
     dbgrids, TofMeth, HCompte, SaisComm, UtilSais, Mul, Vierge, Fe_Main;

Type
  TOF_SAISECART = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    e_general, e_auxiliaire : THEdit;
    e_devise : THValCombobox;
    e_datecomptable : THEdit;
    e_refinterne : THEdit;
    e_libelle : THEdit;
    e_debit : THNumEdit;
    e_credit : THNumEdit;

    BValider : TToolButton;

    aTob: TOB;
    EcartJal, EcartCptD, EcartCptC: string;
    NumPiece: integer;
    Debit, Credit: double;
    memPivot, memOppos : RDevise ;

    procedure AddDeviseItems;
    procedure GeneralExit(Sender: TObject);
    procedure AuxiliaireExit(Sender: TObject);
    procedure DatecomptableExit(Sender: TObject);
    procedure DebitExit(Sender: TObject);
    procedure CreditExit(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure EcritureSaisie(EcrSaisie: tob);
    procedure EcritureGenere(EcrGenere: tob);
    function  GetLastNumeroPiece(DateComptable: TDatetime): integer;
    procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function  MajComptes : Boolean ;
    function  MajJournal : Boolean ;
  end ;

Type
  TOF_LCOMPTE = Class (TOF_Meth)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    FListe : TDBGrid;
    checkbox1 : TCheckbox;
    Exo : THValComboBox ;
    Date1, Date2 : THEdit;
    DateD, DateF : TDatetime;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure OnDblClickFListe(Sender: TObject) ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  uLibAnalytique ;

procedure TOF_SAISECART.OnNew ;
begin
  Inherited ;
  OnDelete;
end ;

procedure TOF_SAISECART.OnDelete ;
begin
  Inherited ;
    e_general.text := '';
    e_auxiliaire.text := '';
    e_devise.ItemIndex := 0;
    e_datecomptable.Text := datetostr(now);
    e_refinterne.text := '';
    e_libelle.text := '';
    SetControlEnabled('e_debit', true);
    SetControlEnabled('e_credit', true);
    SetControlProperty('e_debit', 'Color', clWindow);
    SetControlProperty('e_credit', 'Color', clWindow);
    e_debit.Value := 0.00;
    e_credit.value := 0.00;
end ;

procedure TOF_SAISECART.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SAISECART.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SAISECART.OnArgument (S : String ) ;
begin
  Inherited ;
    e_general := THEdit(GetControl('e_general'));
    e_auxiliaire := THEdit(GetControl('e_auxiliaire'));
    e_devise := THValCombobox(GetControl('e_devise'));
    e_datecomptable := THEdit(GetControl('e_datecomptable'));
    e_refinterne := THEdit(GetControl('e_refinterne'));
    e_libelle := THEdit(GetControl('e_libelle'));
    e_debit := THNumEdit(GetControl('e_debit'));
    e_credit := THNumEdit(GetControl('e_credit'));

    BValider := TToolButton(GetControl('BValider'));

    if e_general <> nil then
       e_general.OnExit := GeneralExit;
    if e_auxiliaire <> nil then
       e_auxiliaire.OnExit := AuxiliaireExit;
    if e_devise <> nil then begin
        AddDeviseItems;
        e_devise.Itemindex := 0;
    end;
    if e_datecomptable <> nil then
       e_datecomptable.OnExit := DatecomptableExit;
    if e_debit <> nil then
       e_debit.OnExit := DebitExit;
    if e_credit <> nil then
       e_credit.OnExit := CreditExit;
    if BValider <> nil then
       BValider.OnClick := BValiderClick;
    if Ecran<>nil then TFVierge(Ecran).OnKeyDown:=OnKeyDown ;
end ;

procedure TOF_SAISECART.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SAISECART.AddDeviseItems;
begin
memPivot.Code:=V_PGI.DevisePivot ;
if memPivot.Code<>'' then GetInfosDevise(memPivot) ;
memOppos.Code:=V_PGI.DeviseFongible ;
if memOppos.Code<>'' then GetInfosDevise(memOppos) ;
e_devise.clear ;
e_devise.items.Add(memPivot.Libelle) ;
e_devise.Values.Add(memPivot.Code) ;
if VH^.TenueEuro
  then begin e_devise.Items.Add(memOppos.Libelle) ; e_devise.Values.Add(memPivot.Code) ; end
  else begin e_devise.Items.Add('Euros') ;          e_devise.Values.Add('EUR') ;         end ;
end;

procedure TOF_SAISECART.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if key=VK_F10 then BValiderClick(nil) ;
end ;

procedure TOF_SAISECART.AuxiliaireExit(Sender: TObject);
begin
if e_auxiliaire.Text<>''
  then e_auxiliaire.Text:=BourreLaDonc(e_auxiliaire.Text, fbAux)
  else begin
  PgiBox('Le compte auxiliaire doit être renseigné','Saisie écart de conversion');
  SetFocusControl('e_auxiliaire');
  end;
end;

procedure TOF_SAISECART.DebitExit(Sender: TObject);
var Val: Double ;
begin
Val:=StrToFloat(e_debit.text);
if Val<>0 then SetControlEnabled('E_CREDIT', FALSE)
          else SetControlEnabled('E_CREDIT', TRUE) ;
end ;

procedure TOF_SAISECART.CreditExit(Sender: TObject);
var Val: Double ;
begin
Val:=StrToFloat(e_credit.text);
if Val<>0 then SetControlEnabled('E_DEBIT', FALSE)
          else SetControlEnabled('E_DEBIT', TRUE) ;
end ;

procedure TOF_SAISECART.DatecomptableExit(Sender: TObject);
begin
  if (strtoDate(e_datecomptable.text) < VH^.Encours.Deb) or (strtoDate(e_datecomptable.text) > VH^.Encours.Fin) then
  begin
     PgiBox('La date comptable n''est pas correcte','Saisie écart de conversion');
     SetFocusControl('e_datecomptable');
  end;
end;

procedure TOF_SAISECART.GeneralExit(Sender: TObject);
var Compte: TGGeneral ;
begin
if e_general.Text<>'' then e_general.Text:=BourreLaDonc(e_general.Text, fbGene) ;
Compte:=TGGeneral.Create(e_general.Text) ;
if Compte.Collectif then SetControlEnabled('E_AUXILIAIRE', TRUE)
                    else begin SetControlEnabled('E_AUXILIAIRE', FALSE) ; e_auxiliaire.text:='' ; end ;
Compte.Free ;
(*
   q := OpenSQL('SELECT  G_SENS FROM GENERAUX WHERE G_GENERAL="' + e_general.Text + '"', true);
   if not q.eof then begin
      if q.fieldbyname('G_SENS').AsString = 'D' then begin
         SetControlEnabled('e_debit', true);
         SetControlEnabled('e_credit', false);
      end
      else if q.fieldbyname('G_SENS').AsString = 'C' then begin
         SetControlEnabled('e_debit', false);
         SetControlEnabled('e_credit', true);
      end
      else begin
         SetControlEnabled('e_debit', true);
         SetControlEnabled('e_credit', true);
      end;
      if e_debit.enabled  then SetControlProperty('e_debit',  'Color', clWindow)
                          else SetControlProperty('e_debit',  'Color', clBtnFace);
      if e_credit.enabled then SetControlProperty('e_credit', 'Color', clWindow)
                          else SetControlProperty('e_credit', 'Color', clBtnFace);
   end;
   ferme(q);
*)   
end;

procedure TOF_SAISECART.BValiderClick(Sender: TObject);
var EcrSaisie, EcrGenere : TOB ; 
begin
   EcartJal   := GetParamSocSecur('SO_JALECARTEURO','');
   NumPiece   := GetLastNumeroPiece(StrToDateTime(e_DateComptable.Text));
   EcartCptD  := GetParamSocSecur('SO_ECCEURODEBIT','');
   EcartCptC  := GetParamSocSecur('SO_ECCEUROCREDIT','');
   Debit      := StrToFloat(e_debit.text);
   Credit     := StrToFloat(e_credit.text);

   aTob := Tob.create('', nil, -1);
   EcrSaisie := Tob.create('ECRITURE', aTob, -1);
   AlloueAxe( EcrSaisie ) ;  // SBO 25/01/2006
   EcrGenere := Tob.create('ECRITURE', aTob, -1);
   AlloueAxe( EcrGenere ) ;  // SBO 25/01/2006
   try
     EcritureSaisie(EcrSaisie);
     EcritureGenere(EcrGenere);
     aTob.InsertDBByNivel(TRUE) ;
     if MajComptes then MajJournal ;
   finally
     EcrSaisie.free;
     EcrGenere.free;
     aTob.free;
   end;
OnNew ;
end;

function TOF_SAISECART.GetLastNumeroPiece(DateComptable: TDatetime) : Integer ;
begin
result:=GetNewNumJal(EcartJal, TRUE, DateComptable) ;
end ;

procedure TOF_SAISECART.EcritureGenere(EcrGenere: tob);
var Compte : TGGeneral ;
begin
  with EcrGenere do begin
       putValue('e_exercice',        VH^.EnCours.Code);
       putValue('e_journal',         EcartJal);
       putValue('e_datecomptable',   strtodate(e_datecomptable.text));
       putValue('e_numeropiece',     NumPiece);
       putValue('e_numligne',        2);
       if Debit <> 0 then
          putValue('e_general',      EcartCptD)
       else
          putValue('e_general',      EcartCptC);
       putValue('e_auxiliaire',      e_auxiliaire.text);
       putValue('e_refinterne',      e_refinterne.text);
       putValue('e_libelle',         e_libelle.text);
       putValue('e_naturepiece',     'OD');
       PutValue('E_QUALIFPIECE',     'N') ;
       PutValue('E_TYPEMVT',         'DIV') ;
       PutValue('E_ETABLISSEMENT',   VH^.EtablisDefaut) ;
       PutValue('E_CONTROLETVA',     'RIE') ;
       PutValue('E_ECRANOUVEAU',     'N') ;
       PutValue('E_PERIODE',         GetPeriode(GetValue('E_DATECOMPTABLE'))) ;
       PutValue('E_SEMAINE',         NumSemaine(GetValue('E_DATECOMPTABLE'))) ;
       if (not VH^.TenueEuro) and (e_devise.ItemIndex=1)
         then PutValue('E_DEVISE',          e_devise.Values[0])
         else PutValue('E_DEVISE',          e_devise.Values[e_devise.ItemIndex]) ;
       if e_devise.ItemIndex=0 then
         begin
         memPivot.DateTaux:=GetValue('E_DATECOMPTABLE') ;
         memPivot.Taux:=GetTaux(memPivot.Code, memPivot.DateTaux, GetValue('E_DATECOMPTABLE')) ;
         PutValue('E_SAISIEEURO',    '-') ;
         PutValue('E_TAUXDEV',       memPivot.Taux) ;
         PutValue('E_DATETAUXDEV',   memPivot.DateTaux) ;
         PutValue('E_DEBIT',         Credit);
         PutValue('E_DEBITDEV',      Credit);
         PutValue('E_CREDIT',        Debit);
         PutValue('E_CREDITDEV',     Debit);
         end else
         begin
         memOppos.DateTaux:=GetValue('E_DATECOMPTABLE') ;
         memOppos.Taux:=GetTaux(memOppos.Code, memOppos.DateTaux, GetValue('E_DATECOMPTABLE')) ;
         PutValue('E_SAISIEEURO',    'X') ;
         PutValue('E_TAUXDEV',       memOppos.Taux) ;
         PutValue('E_DATETAUXDEV',   memOppos.DateTaux) ;
         PutValue('E_DEBITEURO',     Credit);
         PutValue('E_CREDITEURO',    Debit);
         end ;
  end;
Compte:=TGGeneral.Create(EcrGenere.GetValue('E_GENERAL')) ;
if Ventilable(Compte, 0) then
  begin
  VentilerTOB(EcrGenere, '', 0, memPivot.Decimale, FALSE) ;
  EcrGenere.PutValue('E_ANA', 'X') ;
  end ;
Compte.Free ;
end;

procedure TOF_SAISECART.EcritureSaisie(EcrSaisie: tob);
var Compte : TGGeneral ;
begin
  with EcrSaisie do begin
       putValue('e_exercice',        VH^.EnCours.Code);
       putValue('e_journal',         EcartJal);
       putValue('e_datecomptable',   strtodate(e_datecomptable.text));
       putValue('e_numeropiece',     NumPiece);
       putValue('e_numligne',        1);
       putValue('e_general',         e_general.text);
       putValue('e_auxiliaire',      e_auxiliaire.text);
       putValue('e_refinterne',      e_refinterne.text);
       putValue('e_libelle',         e_libelle.text);
       putValue('e_naturepiece',     'OD');
       PutValue('E_QUALIFPIECE',     'N') ;
       PutValue('E_TYPEMVT',         'DIV') ;
       PutValue('E_ETABLISSEMENT',   VH^.EtablisDefaut) ;
       PutValue('E_CONTROLETVA',     'RIE') ;
       PutValue('E_ECRANOUVEAU',     'N') ;
       PutValue('E_PERIODE',         GetPeriode(GetValue('E_DATECOMPTABLE'))) ;
       PutValue('E_SEMAINE',         NumSemaine(GetValue('E_DATECOMPTABLE'))) ;
       if (not VH^.TenueEuro) and (e_devise.ItemIndex=1)
         then PutValue('E_DEVISE',          e_devise.Values[0])
         else PutValue('E_DEVISE',          e_devise.Values[e_devise.ItemIndex]) ;
       if e_devise.ItemIndex=0 then
         begin
         memPivot.DateTaux:=GetValue('E_DATECOMPTABLE') ;
         memPivot.Taux:=GetTaux(memPivot.Code, memPivot.DateTaux, GetValue('E_DATECOMPTABLE')) ;
         PutValue('E_SAISIEEURO',    '-') ;
         PutValue('E_TAUXDEV',       memPivot.Taux) ;
         PutValue('E_DATETAUXDEV',   memPivot.DateTaux) ;
         PutValue('E_DEBIT',         Debit);
         PutValue('E_DEBITDEV',      Debit);
         PutValue('E_CREDIT',        Credit);
         PutValue('E_CREDITDEV',        Credit);
         end else
         begin
         memOppos.DateTaux:=GetValue('E_DATECOMPTABLE') ;
         memOppos.Taux:=GetTaux(memOppos.Code, memOppos.DateTaux, GetValue('E_DATECOMPTABLE')) ;
         PutValue('E_SAISIEEURO',    'X') ;
         PutValue('E_TAUXDEV',       memOppos.Taux) ;
         PutValue('E_DATETAUXDEV',   memOppos.DateTaux) ;
         PutValue('E_DEBITEURO',     Debit);
         PutValue('E_CREDITEURO',    Credit);
         end ;
  end;
Compte:=TGGeneral.Create(EcrSaisie.GetValue('E_GENERAL')) ;
if Ventilable(Compte, 0) then
  begin
  VentilerTOB(EcrSaisie, '', 0, memPivot.Decimale, FALSE) ;
  EcrSaisie.PutValue('E_ANA', 'X') ;
  end ;
Compte.Free ;
end;

function TOF_SAISECART.MajComptes : Boolean ;
var e, ee      : TOB ;
    Compte     : TGGeneral ;
    Aux        : TGTiers ;
    i, j, k, c : Integer ;
    FRM        : TFRM ;
    ll         : LongInt ;
    XD,XC      : Double ;
begin

  Result    := True ;

  // Généraux
  for i:=0 to aTob.Detail.Count-1 do
    begin
    e  := aTob.Detail[i] ;
    XD := e.GetValue('E_DEBIT') ;
    XC := e.GetValue('E_CREDIT') ;

    FRM.Cpt   := e.GetValue('E_GENERAL') ;
    FRM.NumD  := e.GetValue('E_NUMEROPIECE') ;
    FRM.DateD := e.GetValue('E_DATECOMPTABLE') ;
    FRM.LigD  := e.GetValue('E_NUMLIGNE') ;

    Compte:=TGGeneral.Create( e.GetValue('E_GENERAL') ) ;
    AttribParamsComp ( FRM, Compte ) ;
    Compte.Free ;

    AttribParamsNew ( FRM, XD, XC, teEncours ) ;
    ll := ExecReqMAJ ( fbGene, False, True, FRM ) ;
    If ll<>1 then
      begin
      V_PGI.IoError := oeSaisie ;
      result        := False ;
      end ;

    end ;

  // Tiers
  for i:=0 to aTob.Detail.Count-1 do
    begin

    e:=aTob.Detail[i] ;
    if e.GetValue('E_AUXILIAIRE')='' then Continue ;

    FRM.Cpt   := e.GetValue('E_AUXILIAIRE') ;
    FRM.NumD  := e.GetValue('E_NUMEROPIECE') ;
    FRM.DateD := e.GetValue('E_DATECOMPTABLE') ;
    FRM.LigD  := e.GetValue('E_NUMLIGNE') ;
    XD        := e.GetValue('E_DEBIT') ;
    XC        := e.GetValue('E_CREDIT') ;

    Aux := TGTiers.Create( e.GetValue('E_AUXILIAIRE') ) ;
    AttribParamsComp ( FRM, Aux ) ;
    Aux.Free ;

    AttribParamsNew ( FRM, XD, XC, teEncours ) ;
    ll := ExecReqMAJ ( fbAux, False, True, FRM ) ;
    If ll<>1 then
      begin
      V_PGI.IoError := oeSaisie ;
      result        := False ;
      end ;

    end ;

  // Sections
  for i:=0 to aTob.Detail.Count-1 do
    begin
    e:=aTob.Detail[i] ;
    if e.GetValue('E_ANA')<>'X' then Continue ;
    for j:=0 to MAXAXE-1 do
      begin
      c:=e.Detail[j].Detail.Count ;
      if c=0 then Continue ;
      for k:=0 to c-1 do
        begin
        ee := e.Detail[j].Detail[k] ;
        XD := ee.GetValue('Y_DEBIT') ;
        XC := ee.GetValue('Y_CREDIT') ;

        FRM.Cpt   := ee.GetValue('Y_SECTION') ;
        FRM.Axe   := ee.GetValue('Y_AXE') ;
        FRM.NumD  := e.GetValue('E_NUMEROPIECE') ;
        FRM.DateD := e.GetValue('E_DATECOMPTABLE') ;
        FRM.LigD  := e.GetValue('E_NUMLIGNE') ;

        AttribParamsNew ( FRM, XD, XC, teEncours ) ;
        ll := ExecReqMAJ ( fbSect, False, False, FRM ) ;
        If ll<>1 then V_PGI.IoError:=oeSaisie ;
        end ;
      end ;
    end ;

end ;

function TOF_SAISECART.MajJournal : Boolean ;
var FRM   : TFRM ;
    ll    : LongInt ;
begin
  Result:=TRUE ;
  FRM.Cpt   := EcartJal ;
  FRM.NumD  := NumPiece ;
  FRM.DateD := StrToDate( E_datecomptable.text ) ;
  AttribParamsNew ( FRM, Debit, Credit, teEncours ) ;
  ll := ExecReqMAJ ( fbJal, False, False, FRM ) ;
  if ll<>1 then
    begin
    Result:=FALSE ;
    V_PGI.IoError:=oeSaisie ;
    end ;
end ;

{ TOF_LCOMPTE }

procedure TOF_LCOMPTE.OnArgument(S: String);
begin
inherited ;
FListe:=TDBGrid(GetControl('FLISTE'));
checkbox1:=TCheckbox(GetControl('checkbox1')) ;
Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
Date1:=THEdit(GetControl('E_DATECOMPTABLE')) ;
Date2:=THEdit(GetControl('E_DATECOMPTABLE_')) ;
if FListe<>nil then FListe.OnDblClick:=OnDblClickFListe ;
if    Exo<>nil then begin Exo.OnChange:=ExoOnChange ; Exo.Value:=VH^.Entree.Code ; end ;
if  Date1<>nil then Date1.OnExit:=DateOnExit ;
if  Date2<>nil then Date2.OnExit:=DateOnExit ;
TFMul(Ecran).Pages.Pages[1].TabVisible:=FALSE ;
TFMul(Ecran).Pages.Pages[2].TabVisible:=FALSE ;
end;

procedure TOF_LCOMPTE.ExoOnChange(Sender: TObject) ;
begin
DoExoToDateOnChange(Exo, Date1, Date2) ;
DateD:=StrToDate(Date1.Text) ;
DateF:=StrToDate(Date2.Text) ;
end ;

procedure TOF_LCOMPTE.DateOnExit(Sender: TObject) ;
begin
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF) ;
end ;

procedure TOF_LCOMPTE.OnClose;
begin
  inherited;
end;

procedure TOF_LCOMPTE.OnDelete;
begin
  inherited;
end;

procedure TOF_LCOMPTE.OnLoad;
var xxGroupBy : THEdit ;
begin
//HAVING ((SUM(E_DEBIT)-SUM(E_CREDIT)<>0) AND (SUM(E_DEBITEURO)-SUM(E_CREDITEURO)=0))
inherited ;
xxGroupBy:=THEdit(GetControl('XX_GROUPBY')) ;
if xxGroupBy=nil then Exit ;
xxGroupBy.Text:='E_GENERAL, E_AUXILIAIRE HAVING ' ;
if checkbox1.Checked
  then xxGroupBy.Text:=xxGroupBy.Text+' ((SUM(E_DEBIT)-SUM(E_CREDIT)=0) AND (SUM(E_DEBITEURO)-SUM(E_CREDITEURO)<>0)) '
  else xxGroupBy.Text:=xxGroupBy.Text+' ((SUM(E_DEBIT)-SUM(E_CREDIT)<>0) AND (SUM(E_DEBITEURO)-SUM(E_CREDITEURO)=0)) '
(*
  Fliste.DataSource.Dataset.Filtered := false;
  inherited;
  if checkbox1.Checked then
     Fliste.DataSource.Dataset.Filter := 'C6=0 AND C7<>0'
  else
     Fliste.DataSource.Dataset.Filter := 'C6<>0 AND C7=0';
  Fliste.DataSource.Dataset.Filtered := true;
*)  
end;

procedure TOF_LCOMPTE.OnNew;
begin
  inherited;
end;

procedure TOF_LCOMPTE.OnUpdate;
begin
  inherited;
end;

procedure TOF_LCOMPTE.OnDblClickFListe(Sender: TObject) ;
begin
AGLLanceFiche('CP', 'ECSAISIE', '', '', '') ;
end;

Initialization
  registerclasses ( [ TOF_SAISECART ] ) ;
  registerclasses ( [ TOF_LCOMPTE ] ) ;
end.
