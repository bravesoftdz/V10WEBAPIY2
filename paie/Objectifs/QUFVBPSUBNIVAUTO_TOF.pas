{***********UNITE*************************************************
Auteur  ...... : EV5
Créé le ...... : 21/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPSUBNIVAUTO ()
Mots clefs ... : TOF;QUFVBPSUBNIVAUTO
*****************************************************************}
Unit QUFVBPSUBNIVAUTO_TOF ;

Interface

Uses UTOF,Graphics,Windows,Controls,Classes;

Type
  TOF_QUFVBPSUBNIVAUTO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    Canvas0: TCanvas;
    Capturing: boolean;
    Captured: boolean;
    StartPlace: TPoint;
    EndPlace: TPoint;
    procedure ClickAll( Sender: TObject );
    procedure ClickNone( Sender: TObject );
    procedure PanelDatesMouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure PanelDatesMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure PanelDatesMouseUp(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
  end ;

function MakeRect(Pt1, Pt2: TPoint; Control: TControl): TRect;
Function TextSize(Phrase : string; Police : TFont = nil) : TPoint;

var NbDates,numnoeud,NodeNiveau:integer;
    LastLevel:boolean;
    FonctionOrig,CodeSession,CodeAxe,LibAxe,ValeurCodeAxe,ValeurLibAxe : string;
    CodeAxePrec,LibAxePrec,ValeurLibAxePrec:string;
    DatesSession,NumDateOK:string ;


Implementation

uses StdCtrls,ExtCtrls,Sysutils,ComCtrls,
     HCtrls,HEnt1,HMsgBox,vierge,HTB97,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$else}eMul,uTob,
     {$ENDIF}
     Uutil,BPBasic,BPUtil,BPFctSession;

procedure TOF_QUFVBPSUBNIVAUTO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnUpdate ;
var ValuesChecked,sublevel:string;
NbDateSelected,i,rep:integer;
okval:boolean;
begin
  Inherited ;
  NbDateSelected:=0;
  for i:= 0 to NbDates-1 do
  begin
    if THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(i))).Checked then
    begin
      if ValuesChecked='' then ValuesChecked:=THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(i))).Caption
      else ValuesChecked:=ValuesChecked + '|' + THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(i))).Caption;
      NbDateSelected:=NbDateSelected+1;
    end;
  end;

  OkVal:=false;
  for i:=1 to 7 do
  begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
    if i <> 7 then
    begin
      if GetControlText('EDTPRCTEVOL'+IntToStr(i)) <> '' then begin OkVal:=true ; break end;
      if GetControlText('EDTVALEVOL'+IntToStr(i)) <> '' then begin OkVal:=true ; break end;
      if GetControlText('EDTSAISIEEVOL'+IntToStr(i)) <> '' then begin OkVal:=true ; break end;
    end
    else
    begin
      if GetControlText('EDTPRCTEVOLQTE') <> '' then begin OkVal:=true ; break end;
      if GetControlText('EDTVALEVOLQTE') <> '' then begin OkVal:=true ; break end;
      if GetControlText('EDTSAISIEEVOLQTE') <> '' then begin OkVal:=true ; break end;
    end
  end;

  if NbDateSelected = 0 then
  begin
    if FonctionOrig='CREATESOUSNIVO' then HShowmessage('1;Création de sous niveaux en série;Aucune Date sélectionnée.;I;O;O;O', '', '')
    else HShowmessage('1;Modification de sous niveaux en série;Aucune Date sélectionnée.;I;O;O;O', '', '');
    SetControlText('EdtFerme','NON');
    TFVierge(ecran).Retour := '';
  end
  else if ((OkVal = false) AND (FonctionOrig='MODIFSOUSNIVO'))then
  begin
    HShowmessage('1;Modification de sous niveaux en série;Aucune valeur sélectionnée.;I;O;O;O', '', '');
    SetControlText('EdtFerme','NON');
    TFVierge(ecran).Retour := '';
  end
  else
  begin
    SetControlText('EdtFerme','OUI');
    if NbDateSelected=1 then sublevel:=TraduireMemoire(' sous niveau') else sublevel:=TraduireMemoire(' sous niveaux');
    if FonctionOrig='CREATESOUSNIVO' then
    begin
      if ValeurLibAxePrec=''
      then rep := HShowmessage('1;Création de sous niveaux en série;Vous allez créer '
                          +IntToStr(NbDateSelected)+sublevel+' pour '+LibAxe+' = '+ValeurLibAxe+
                          #13#10 +' Etes-vous sûr ?;Q;YN;N;N', '', '')
      else rep := HShowmessage('1;Création de sous niveaux en série;Vous allez créer '
                          +IntToStr(NbDateSelected)+sublevel+' pour '+LibAxe+' = '+ValeurLibAxe+#13#10+
                          ' et '+LibAxePrec+' = '+ValeurLibAxePrec + #13#10 +' Etes-vous sûr ?;Q;YN;N;N', '', '')
    end
    else
    begin
      if LastLevel
      then rep := HShowmessage('1;Modification de sous niveaux en série;Vous allez modifier '
                               +IntToStr(NbDateSelected)+' valeurs dates pour '+LibAxe+' = '+ValeurLibAxe
                               +#13#10+' Etes-vous sûr ?;Q;YN;N;N', '', '')
      else rep := HShowmessage('1;Modification de sous niveaux en série;Vous allez modifier '
                               +IntToStr(NbDateSelected)+' valeurs dates pour tous les'+sublevel+#13#10
                               +' de '+LibAxe+' = '+ValeurLibAxe+#13#10+' Etes-vous sûr ?;Q;YN;N;N', '', '');
    end;
    if rep = mrYes
    then TFVierge(ecran).Retour:='DATESCHECKED='+ValuesChecked+';EVOLPRCT1='+GetControlText('EDTPRCTEVOL1')+
    ';EVOLVAL1='+GetControlText('EDTVALEVOL1')+';SAISIE1='+GetControlText('EDTSAISIEEVOL1')+
    ';EVOLPRCT2='+GetControlText('EDTPRCTEVOL2')+';EVOLVAL2='+GetControlText('EDTVALEVOL2')+
    ';SAISIE2='+GetControlText('EDTSAISIEEVOL2')+';EVOLPRCT3='+GetControlText('EDTPRCTEVOL3')+
    ';EVOLVAL3='+GetControlText('EDTVALEVOL3')+';SAISIE3='+GetControlText('EDTSAISIEEVOL3')+
    ';EVOLPRCT4='+GetControlText('EDTPRCTEVOL4')+';EVOLVAL4='+GetControlText('EDTVALEVOL4')+
    ';SAISIE4='+GetControlText('EDTSAISIEEVOL4')+';EVOLPRCT5='+GetControlText('EDTPRCTEVOL5')+
    ';EVOLVAL5='+GetControlText('EDTVALEVOL5')+';SAISIE5='+GetControlText('EDTSAISIEEVOL5')+
    ';EVOLPRCT6='+GetControlText('EDTPRCTEVOL6')+';EVOLVAL6='+GetControlText('EDTVALEVOL6')+
    ';SAISIE6='+GetControlText('EDTSAISIEEVOL6')+';EVOLPRCTQTE='+GetControlText('EDTPRCTEVOLQTE')+
    ';EVOLVALQTE='+GetControlText('EDTVALEVOLQTE')+';SAISIEQTE='+GetControlText('EDTSAISIEEVOLQTE')
    else
    begin
      TFVierge(ecran).Retour := '';
      SetControlText('EdtFerme','NON');
    end;
  end;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnLoad ;
var Q:TQuery;
i,numDate:integer;
DateFound:boolean;
DatesOfSession:string;
DateTmp:TDateTime;
NivPrec:string;
begin
  Inherited ;
  if FonctionOrig='CREATESOUSNIVO' then
  begin
    //Création de sous-niveaux Dates
    numDate:=0;
    DateFound:=false;
    if NodeNiveau=1 then NivPrec := 'QBR_VALAXENIV1'
    else NivPrec := 'QBR_VALAXENIV'+IntToStr(NodeNiveau-1);
    Q:=OpenSQL('SELECT QBR_VALEURAXE,'+NivPrec+' FROM QBPARBRE WHERE QBR_CODESESSION="'+CodeSession+'" AND QBR_NUMNOEUDPERE='+IntToStr(numnoeud)+
               ' ORDER BY QBR_VALEURAXE',true);
    While not Q.Eof do
    begin
      DatesOfSession:=DatesSession;
      i:=0 ;

      While DatesOfSession<>'' do
      begin
        DateTmp:=StrToDateTime(ReadTokenSt(DatesOfSession));
        if DateTmp=StrToDateTime(Q.Fields[0].AsString) then
        begin
          DateFound:=true;
          numDate:=i;
          if NumDateOK='' then NumDateOK:=IntToStr(numDate)
          else NumDateOK:=NumDateOK+';'+IntToStr(numDate);
          break;
        end;
        i:=i+1;
      end;

      if DateFound then
      begin
        THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(numDate))).Enabled := false;
        THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(numDate))).State := cbGrayed;
      end;
      DateFound:=false;
      Q.Next;
    end;
    ValeurLibAxePrec := DonneLibelleValeurAxe(CodeAxePrec,Q.Fields[1].AsString);
    Ferme(Q);

    if ((ValeurLibAxePrec='') AND (NodeNiveau<>1)) then
    begin
      Q:=OpenSQL('SELECT '+NivPrec+' FROM QBPARBRE WHERE QBR_CODESESSION="'+CodeSession+'" AND QBR_NUMNOEUD='+IntToStr(numnoeud),true);
      if not Q.Eof then ValeurLibAxePrec := DonneLibelleValeurAxe(CodeAxePrec,Q.Fields[0].AsString);
      Ferme(Q);
    end;
  end;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnArgument (S : String ) ;
var CBDynamique:THCheckBox;
DateDeb,DateFin,DateCur:TDateTime;
Lignes,Colonnes,i,ValeurAffiche,NbValAff,Onglet,OngletSav:integer;
TabValAff,TabLibelle:array [0..7] of hString;
Cadre : TPoint;
begin
  Inherited ;
  StartPlace := Point(-1, -1);
  CodeSession:=TrouveArgument(S,'SESSION','');
  ValeurAffiche:=DonneValeurAffiche(codesession);
  for i:=1 to 7 do THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=false;
  Case ContextBP of
    0,1 : begin //Mode-GC
          case ValeurAffiche of
            1 : begin
                  TabValAff[0]:='PTC';
                  THTabSheet(GetControl('TAB1')).tabVisible:=true;
                end;
            2 : begin
                  TabValAff[0]:='QTE';
                  THTabSheet(GetControl('TAB2')).tabVisible:=true;
                end;
             3 : begin
                  TabValAff[0]:='PHT';
                  THTabSheet(GetControl('TAB3')).tabVisible:=true;
                end;
            4 : begin
                  TabValAff[0]:='UTC';
                  THTabSheet(GetControl('TAB4')).tabVisible:=true;
                end;
            5 : begin
                  TabValAff[0]:='UHT';
                  THTabSheet(GetControl('TAB5')).tabVisible:=true;
                end;
            6 : begin
                  TabValAff[0]:='PAH';
                  THTabSheet(GetControl('TAB6')).tabVisible:=true;
                end;
            7 : begin
                  TabValAff[0]:='MAR';
                  THTabSheet(GetControl('TAB7')).tabVisible:=true;
                end
          end
        end;
    2 : begin //Compta
          case ValeurAffiche of
            1 : begin
                  TabValAff[0]:='DC1';
                  THTabSheet(GetControl('TAB1')).tabVisible:=true;
                  THTabSheet(GetControl('TAB1')).Caption := 'Débit - Crédit';
                end;
            3 : begin
                  TabValAff[0]:='CD1';
                  THTabSheet(GetControl('TAB3')).tabVisible:=true;
                  THTabSheet(GetControl('TAB3')).Caption := 'Crédit - Débit';
                end;
          end
        end;
    3 : begin //Paie
          LibValAff(codeSession,TabLibelle);
          NbValAff := StrToInt(TabLibelle[0]);
          for i:=1 to NbValAff do
          begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
            if i = 1 then THTabSheet(GetControl('TAB2')).tabVisible := true
            else if i = 2 then THTabSheet(GetControl('TAB1')).tabVisible := true
            else THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible := true;
            if i = 1 then THTabSheet(GetControl('TAB2')).Caption := TabLibelle[i]
            else if i = 2 then THTabSheet(GetControl('TAB1')).Caption := TabLibelle[i]
            else THTabSheet(GetControl('TAB'+IntToStr(i))).Caption := TabLibelle[i]
          end;
          if NbValAff > 1 then
          begin
            TPageControl(getControl( 'PAGE')).Pages[1].PageIndex := 0;
            TPageControl(getControl( 'PAGE')).ActivePage := THTabSheet(GetControl('TAB2'));
          end;
        end
  end; //CASE

  TToolBarButton97(GetControl('SELECTALL')).onClick := ClickAll;
  TToolBarButton97(GetControl('SELECTNONE')).onClick := ClickNone;
  numnoeud:=StrToInt(TrouveArgument(S,'NUMNOEUD',''));
  DateDeb:=StrToDateTime(TrouveArgument(S,'DATEDEB',''));
  DateFin:=StrtoDateTime(TrouveArgument(S,'DATEFIN',''));
  NodeNiveau:=StrToInt(TrouveArgument(S,'NIVEAU',''));
  FonctionOrig:=TrouveArgument(S,'FONCTION','');
  if NodeNiveau=ChercheNivMaxSession(CodeSession) then LastLevel:=true else LastLevel:=false;

  if FonctionOrig='CREATESOUSNIVO' then
  begin
    THEdit(GetControl('EDTPRCTEVOL1')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOL1')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL1')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOL2')).Enabled := false;
    THEdit(GetControl('EDTVALEVOL2')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOL3')).Enabled := false;
    THEdit(GetControl('EDTVALEVOL3')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOL4')).Enabled := false;
    THEdit(GetControl('EDTVALEVOL4')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOL5')).Enabled := false;
    THEdit(GetControl('EDTVALEVOL5')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOL6')).Enabled := false;
    THEdit(GetControl('EDTVALEVOL6')).Enabled := false;
    THEdit(GetControl('EDTPRCTEVOLQTE')).Enabled := false;
    THEdit(GetControl('EDTVALEVOLQTE')).Enabled := false;

    THEdit(GetControl('EDTPRCTEVOL1')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL1')).Color  := clBtnFace;
    THEdit(GetControl('EDTPRCTEVOL2')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL2')).Color  := clBtnFace;
    THEdit(GetControl('EDTPRCTEVOL3')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL3')).Color  := clBtnFace;
    THEdit(GetControl('EDTPRCTEVOL4')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL4')).Color  := clBtnFace;
    THEdit(GetControl('EDTPRCTEVOL5')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL5')).Color  := clBtnFace;
    THEdit(GetControl('EDTPRCTEVOL6')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOL6')).Color  := clBtnFace;
    THEdit(GetControl('EDTPRCTEVOLQTE')).Color  := clBtnFace;
    THEdit(GetControl('EDTVALEVOLQTE')).Color  := clBtnFace;
  end;

  NbDates:=0;
  Lignes:=0;
  Colonnes:=0;
  DatesSession:='';
  NumDateOK:='';
  DateCur:=DateDeb;
  //Cache les tabs
  for Onglet := 0 to TPageControl(GetControl('PCDATES')).PageCount - 1 do
  begin
    TPageControl(GetControl('PCDATES')).Pages[Onglet].TabVisible := false;
  end;
  TPageControl(GetControl('PCDATES')).ActivePageIndex := 0;
  Onglet:=0;
  OngletSav:=0;

  While DateCur<DateFin do
  begin
    if DatesSession='' then DatesSession:=DateTimeToStr(DateCur)
    else DatesSession:=DatesSession+';'+DateTimeToStr(DateCur);
    if OngletSav<>Onglet then
    begin
      if Onglet = 10 then
      begin
        //Nombre de périodes limites = 10*48 = 480
        HShowMessage('1;Modification de sous niveaux en série;Il n''est pas possible de gérer plus de #13#10'+
                     ' 480 périodes avec cet outil.;W;O;O;O;', '', '');
        exit;
      end;
      if Onglet = 1 then TPageControl(GetControl('PCDATES')).Pages[0].TabVisible := true;
      TPageControl(GetControl('PCDATES')).Pages[Onglet].TabVisible := true;
      OngletSav:=Onglet;
    end;
    CBDynamique := THCheckBox.Create(TheMainForm );
    CBDynamique.Font.Name := 'Ms Sans Serif';
    CBDynamique.Font.Height := 8;
    CBDynamique.Font.Style := [];
    CBDynamique.Parent := TPanel(GetControl('PDATES'+IntToStr(Onglet)));
    CBDynamique.Top := 4 + Lignes * 20;
    CBDynamique.Left := 8 + (Colonnes * 100);
    CBDynamique.Name := 'CBDate'+IntToStr(NbDates);
    CBDynamique.Caption := DateTimeToStr(DateCur);
    Cadre := TextSize(DateTimeToStr(DateCur),CBDynamique.Font);
    CBDynamique.Height := Cadre.Y;
    CBDynamique.Width := Cadre.X + 16; //Largeur du checkbox
    NbDates:=NbDates+1;
    Lignes:=Lignes+1;
    if Lignes = 12 then
    begin
      Lignes := 0;
      Colonnes := Colonnes + 1;
    end;
    if Colonnes = 4 then
    begin
      Lignes := 0;
      Colonnes := 0;
      OngletSav:= Onglet;
      Onglet := Onglet + 1;
    end;
    DateCur:=DateCur+32;
    DateCur:=DEBUTDEMOIS(DateCur);
  end;

  for i:=0 to Onglet do
  begin
    if i=10 then break;
    TPanel(GetControl('PDATES'+IntToStr(i))).OnMouseDown := PanelDatesMouseDown;
    TPanel(GetControl('PDATES'+IntToStr(i))).OnMouseMove := PanelDatesMouseMove;
    TPanel(GetControl('PDATES'+IntToStr(i))).OnMouseUp := PanelDatesMouseUp;
  end;

  CodeAxe := TrouveArgument(S,'CODEAXE','');
  LibAxe := DonneLibelleAxe(CodeAxe);
  ValeurCodeAxe := TrouveArgument(S,'VALEURCODEAXE','');
  ValeurLibAxe := DonneLibelleValeurAxe(CodeAxe,ValeurCodeAxe);

  CodeAxePrec := TrouveArgument(S,'CODEAXEPREC','');
  LibAxePrec := DonneLibelleAxe(CodeAxePrec);

  if FonctionOrig='CREATESOUSNIVO' then Ecran.Caption :=TraduireMemoire('Création : '+Libaxe+' = '+ValeurLibAxe)
  else Ecran.Caption :=TraduireMemoire('Modification : '+Libaxe+' = '+ValeurLibAxe);
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnClose ;
var i:integer;
begin
  Inherited ;
  if GetControlText('EdtFerme')='OUI' then
  begin
    for i:= 0 to NbDates-1 do THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(i))).Free;
  end;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPSUBNIVAUTO.ClickAll( Sender: TObject );
var i,NumOK:integer;
NumDatesTmp:string;
Modif:boolean;
begin
  NumDatesTmp:=NumDateOK;
  for i:= 0 to NbDates-1 do
  begin
    Modif:=true;
    while NumDatesTmp<>'' do
    begin
      NumOK:=StrToInt(ReadTokenSt(NumDatesTmp));
      if i=NumOK then
      begin
        Modif:=false;
        break;
      end;
    end;
    if Modif then THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(i))).State := cbChecked;
  end;
end;

procedure TOF_QUFVBPSUBNIVAUTO.ClickNone( Sender: TObject );
var i,NumOK:integer;
NumDatesTmp:string;
Modif:boolean;
begin
  NumDatesTmp:=NumDateOK;
  for i:= 0 to NbDates-1 do
  begin
    Modif:=true;
    while NumDatesTmp<>'' do
    begin
      NumOK:=StrToInt(ReadTokenSt(NumDatesTmp));
      if i=NumOK then
      begin
        Modif:=false;
        break;
      end;
    end;
    if Modif then THCheckBox(TheMainForm.FindComponent('CBDate'+intToStr(i))).State := cbunChecked;
  end;
end;

procedure TOF_QUFVBPSUBNIVAUTO.PanelDatesMouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
var CurrentPanel:TPanel;
begin
  CurrentPanel :=  TPanel(GetControl(TControl( Sender ).name));
  Canvas0 := TCanvas.Create;
  Canvas0.Handle := GetDc(0);
  StartPlace := CurrentPanel.ClientToScreen(Point(X, Y));
  EndPlace := StartPlace;
  DrawFocusRect(Canvas0.Handle, MakeRect(StartPlace, EndPlace,CurrentPanel));
  Capturing := true;
  Captured := true;
end;

procedure TOF_QUFVBPSUBNIVAUTO.PanelDatesMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var OutRect,Rect1,Rect2: TRect;
    i: integer;
    CurrentPanel: TPanel;
    SelCheckBox:TCheckBox;
begin
  CurrentPanel :=  TPanel(GetControl(TControl( Sender ).name));
  if Capturing then
  begin
    DrawFocusRect(Canvas0.Handle, MakeRect(StartPlace, EndPlace, CurrentPanel));
    EndPlace := CurrentPanel.ClientToScreen(Point(X, Y));
    for i := 0 to CurrentPanel.ControlCount - 1 do
      if CurrentPanel.Controls[i] is TCheckbox then
      begin
        SelCheckBox := CurrentPanel.Controls[i] as TCheckbox;
        Rect1 := SelCheckBox.BoundsRect;
        Rect1.TopLeft := SelCheckBox.ClientToScreen(Point(0,0));
        Rect1.BottomRight := SelCheckBox.ClientToScreen(Point(SelCheckBox.Width, SelCheckBox.Height));
        Rect2 := MakeRect(StartPlace, EndPlace, CurrentPanel);

        if Shift=[ssCtrl..ssLeft] then
        begin
          if SelCheckBox.Enabled then
            if IntersectRect(OutRect, Rect1, Rect2) then SelCheckBox.Checked := false
        end
        else
        begin
          if SelCheckBox.Enabled then
            if IntersectRect(OutRect, Rect1, Rect2) then SelCheckBox.Checked := true;
        end;
         SelCheckBox.Update;
      end;
    DrawFocusRect(Canvas0.Handle, MakeRect(StartPlace, EndPlace, CurrentPanel));
  end;
end;

procedure TOF_QUFVBPSUBNIVAUTO.PanelDatesMouseUp(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
var CurrentPanel : TPanel;
begin
  CurrentPanel :=  TPanel(GetControl(TControl( Sender ).name));
  if Captured then DrawFocusRect(Canvas0.Handle, MakeRect(StartPlace, EndPlace, CurrentPanel));
  ReleaseDC(0, Canvas0.Handle);
  Canvas0.Free;
  Capturing := false;
end;

function MakeRect(Pt1, Pt2: TPoint; Control: TControl): TRect;
var Bounds: TRect;
begin
  if pt1.x < pt2.x then
  begin
    Result.Left := pt1.x;
    Result.Right := pt2.x;
  end
  else
  begin
    Result.Left := pt2.x;
    Result.Right := pt1.x;
  end;
  if pt1.y < pt2.y then
  begin
    Result.Top := pt1.y;
    Result.Bottom := pt2.y;
  end
  else
  begin
    Result.Top := pt2.y;
    Result.Bottom := pt1.y;
  end;

  Bounds.TopLeft := Control.ClientToScreen(Point(0, 0));
  Bounds.BottomRight := Control.ClientToScreen(Point(Control.Width, Control.Height));
  if Result.Left < Bounds.Left then Result.Left := Bounds.Left;
  if Result.Top < Bounds.Top then Result.Top := Bounds.Top;
  if Result.Right > Bounds.Right then Result.Right := Bounds.Right;
  if Result.Bottom > Bounds.Bottom then Result.Bottom := Bounds.Bottom;
end;

Function TextSize(Phrase : string; Police : TFont = nil) : TPoint;
var C: Graphics.TBitmap ;
    DC: HDC;
    Rect: TRect;
begin
  c := Graphics.TBitmap.create;
  if police <> nil then  C.canvas.Font := police;
  Rect.Left := 0;
  Rect.Top:=0;
  Rect.Right:=0;
  Rect.Bottom:=0;
  DC := GetDC(0);
  C.Canvas.Handle := DC;
  DrawText(C.Canvas.Handle, PChar(Phrase), Length(Phrase), Rect, (DT_EXPANDTABS or DT_CALCRECT));
  C.Canvas.Handle := 0;
  ReleaseDC(0, DC);
  result.X:=Rect.Right-Rect.Left;
  result.Y:=Rect.Bottom-Rect.Top;
  C.Free;
end;

Initialization
  registerclasses ( [ TOF_QUFVBPSUBNIVAUTO ] ) ;
end.

