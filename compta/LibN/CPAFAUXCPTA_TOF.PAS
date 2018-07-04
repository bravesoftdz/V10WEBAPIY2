{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPAFAUXCPTA ()
Mots clefs ... : TOF;CPAFAUXCPTA
*****************************************************************}
Unit CPAFAUXCPTA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     sysutils, ComCtrls,
     HCtrls,  // Contrôles Halley
     HDB,     // Contrôles Halley DB
     HEnt1,   // EnableControls,TraduireMemoire
     AGLInit, // TheTob
     HMsgBox, // PGIBox,PGIAsk
     Utob,    // TOB
     Mul,     // TFMul
     Ent1,    // ExoToDates
     FE_Main, // AGLLanceFiche
     UTOF     // TOF
     ;

Procedure TLAuxversTLEcr ;

Type
  TOF_CPAFAUXCPTA = Class (TOF)
  private
    TobTL, TobTabLib, TobEcr : Tob ;
    Fiche : TFMul ;
    QlisteEcr : TQuery ;
    Lecr : TStringList ;
    gszTBL     : String ;
    gszZoneECR : String;
    gbAvecEcrasement, gbAvecLesAnas, gbAttention : Boolean ;
    procedure ExerciceOnClick(Sender: TObject);
    procedure BOuvrirOnClick(Sender: TObject);
    procedure LibereTobs;
    Function  RechercheValeurTableLibre(Aux: String) : String ;
    procedure ChargeLignesAModifier;
    procedure AjouteLesAnas;
    procedure ChargeLecriture(i: Integer; var OkOk: Boolean);  // FQ 10951
    procedure RechercheLesBornes;
    procedure LanceLeTraitement(Deb, Fin: Integer);
    procedure ChargeTobTableLibreDesTiers;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;


Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  CPAFAUXCPTAPARAM_TOF;  // TLAuxversTLEcrParam

Procedure TLAuxversTLEcr ;
BEGIN
AGLLanceFiche ('CP','CPAFAUXCPTA','','','');
END ;

procedure TOF_CPAFAUXCPTA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTA.OnLoad ;
begin
  Inherited ;
end ;

// Rempli les dates en fonction de l'exercice  // FQ 10951
procedure TOF_CPAFAUXCPTA.ExerciceOnClick (Sender : TObject) ;
begin
  ExoToDates(GetControlText('E_Exercice'),GetControl('E_DATECOMPTABLE',True),GetControl('E_DATECOMPTABLE_',True));
end ;

procedure TOF_CPAFAUXCPTA.ChargeTobTableLibreDesTiers ;
Var Q : TQuery  ;
begin
TobTabLib:=Tob.create('TLAux',Nil,-1) ;
Q:=OpenSql('SELECT T_AUXILIAIRE,NT_NATURE FROM TIERS LEFT JOIN NATCPTE ON NT_NATURE=T_TABLE'+copy(gszTBL,3,1)+' WHERE NT_TYPECPTE="T0'+copy(gszTBL,3,1)+'"',True) ;
If Not Q.EOF then TobTabLib.LoadDetailDB('TLAux','','',Q,False) ;
Ferme(Q) ;
end ;

 // FQ 10951
procedure TOF_CPAFAUXCPTA.BOuvrirOnClick (Sender : TObject) ;
var i : Integer ;
    OkOk, ModeBordereau : Boolean ;
    Fiche : TFMul ;
    ListeEcr : THDBGrid ;    
BEGIN
EnableControls(Ecran, False ) ;
Fiche := TFMul(Ecran) ;
ListeEcr := Fiche.FListe ;
QListeEcr := Fiche.Q ;

// Si pas de sélection
If (not ListeEcr.AllSelected) and (ListeEcr.NbSelected=0) then begin
  EnableControls(Ecran, True );
  PgiBox (TraduireMemoire('Aucune ligne sélectionnée. Traitement impossible.'), TraduireMemoire(Ecran.Caption)) ; Exit ;
  Exit;
end;

TLAuxversTLEcrParam;
if TheTOB = nil then begin EnableControls(Ecran, True ); Exit; end;
  gszTBL := TheTOB.GetValue('TBL');
  gszZoneECR := TheTOB.GetValue('ZONEECR');
  gbAvecLesAnas := TheTOB.GetValue('ECRANA');
  gbAvecEcrasement := TheTOB.GetValue('ECRASER');
FreeAndNil(TheTob);

If PgiAsk(TraduireMemoire('Confirmez-vous le traitement ?'),TraduireMemoire(Ecran.Caption))=MrNo then begin EnableControls(Ecran, True); exit; end; // 14603

ChargeLignesAModifier ;
If Lecr.Count<=0 Then begin EnableControls(Ecran, True); Exit ; end;
Fiche.Q.close ;
  gbAttention:=False ;
ChargeTobTableLibreDesTiers ;

Try
  BeginTrans ;
  TobEcr:=Tob.Create('_Ecriture', Nil, -1) ;
    For i := 0 to LEcr.Count-1 do BEGIN
    OkOk:= True ;
    ChargeLecriture(i,OkOk) ; if not OkOk then continue ;
    ModeBordereau:=(TobEcr.Detail[0].GetValue('E_MODESAISIE')='BOR') ;
    if ModeBordereau then RechercheLesBornes else LanceLeTraitement(0, TobEcr.Detail.Count-1) ;
    Committrans ;
    end ;
Except
  RollBack ;
  PgiBox (TraduireMemoire('Fin anormale du traitement.'), TraduireMemoire(Ecran.Caption)) ;
  Liberetobs ;
  EnableControls(Ecran, True) ;
  exit ;
end ;
Fiche.Q.Open ;
  if gbAttention then PgiBox(TraduireMemoire('ATTENTION : certaines écritures n''ont pu être traîtées.'), TraduireMemoire(Ecran.Caption)) ;
PgiInfo (TraduireMemoire('Fin du traitement.'), TraduireMemoire(Ecran.Caption)) ;
Liberetobs ;
EnableControls(Ecran, True) ;
TButton(GetControl('BCHERCHE',True)).Click ;
end ;

procedure TOF_CPAFAUXCPTA.RechercheLesBornes ;
var Total : Double ;
    x,w : Integer ;
begin
Total:=0 ; w:=0 ;
for x:=0 to TobEcr.Detail.Count-1 do
  begin
  Total:=arrondi(Total+TobEcr.detail[x].GetValue('E_DEBIT')-TobEcr.detail[x].GetValue('E_CREDIT'),4) ;
  If Total=0 then begin LanceLeTraitement(w, x) ; Total:=0 ; w:=x+1 ; end ;
  end ;
end ;

procedure TOF_CPAFAUXCPTA.LanceLeTraitement(Deb,Fin : Integer) ;
Var Laux, Lib, Lauxiliaire : String ;
    j,m : Integer ;
    TobA :Tob ;
Begin
Lauxiliaire:='' ; Laux:='' ;
for j:=deb to fin do
  begin
  Laux:=Tobecr.detail[j].GetValue('E_AUXILIAIRE') ;
  if Laux<>'' then
    if Lauxiliaire='' then Lauxiliaire:=Laux
                      else if Lauxiliaire <> Laux then begin gbAttention:=True ; exit ; end ;
  end ;

Lib:=RechercheValeurTableLibre(Lauxiliaire) ;

If Lib<>'' then
  for j:=deb to fin do
    begin
    if (TobEcr.Detail[j].GetValue('E_'+gszZoneECR)='')
    or ((TobEcr.Detail[j].GetValue('E_'+gszZoneECR)<>'') and (gbAvecEcrasement)) then
      begin
      TobEcr.Detail[j].PutValue('E_'+gszZoneECR, lib) ;
      if ((TobEcr.Detail[j].GetValue('E_ANA')='X') and (gbAvecLesAnas)) then
        Begin
        TobA:=TobEcr.Detail[j] ;
        for m:=0 to TobA.Detail.Count-1 do
          begin
          TobA.Detail[m].PutValue('Y_'+gszZoneECR, Lib);
          end ;
        end ;
      end ;
    end ;
TobEcr.UpdateDB ;
end ;

Function TOF_CPAFAUXCPTA.RechercheValeurTableLibre(Aux: String) : String ;
var TobA : Tob ;
begin
Result :='' ;
TobA:=ToBTabLib.FindFirst(['T_AUXILIAIRE'],[Aux],False) ;
If TobA<>Nil then Result:=TobA.GetValue('NT_NATURE')
end ;

procedure TOF_CPAFAUXCPTA.ChargeLecriture(i : Integer ; var OkOk : Boolean) ; // FQ 10951
Var St, Sx, StJal, StExo, StQual : String ;
    StDate: TdateTime ;
    StNum  : Integer ;
    Q: Tquery ;
begin
St:=LEcr[i] ; StJal:='' ; StExo :='' ; StQual :='' ; StDate:=0 ; StNum :=0 ;
If St<>'' Then StExo:=ReadToKenST(St) ;
If St<>'' Then StJal:=Trim(ReadToKenST(St)); // FQ 10951
If St<>'' Then StQual:=ReadToKenST(St) ;
If St<>'' Then StDate:=StrToDate(ReadToKenST(St)) ;
If St<>'' Then StNum:=StrToInt(ReadToKenST(St)) ;

If (StJal='') Or (StExo='') or (StQual='') or (StDate=0) or (StNum=0) Then begin OkOk := False ; exit ; end ;
TobEcr.ClearDetail ;
Sx := 'SELECT E_JOURNAL, E_GENERAL, E_AUXILIAIRE, E_EXERCICE, E_DATECOMPTABLE, E_MODESAISIE, E_NUMEROPIECE, E_NUMLIGNE, '
    + 'E_NUMECHE, E_QUALIFPIECE, E_ANA, E_DEBIT, E_CREDIT, E_'+gszZoneECR+' FROM ECRITURE WHERE '
    + 'E_EXERCICE="'+StExo+'" AND '
    + 'E_JOURNAL="'+StJal+'" AND '
    + 'E_DATECOMPTABLE="'+UsDateTime(StDate)+'" AND '
    + 'E_QUALIFPIECE="'+StQual+'" AND '
    + 'E_NUMEROPIECE="'+IntToStr(StNum)+'"' ;
Q := OpenSql(Sx, True) ;
TobEcr.LoadDetailDB('ECRITURE','','',Q,False,True) ;
Ferme(Q) ;
If gbAvecLesAnas then AjouteLesAnas ;
end ;

procedure TOF_CPAFAUXCPTA.AjouteLesAnas ;
Var x, y : integer ;
    TOBL, TobAna : TOB ;
    SQL : String ;
    Q   : TQuery ;
BEGIN
TobAna:=TOB.Create('',Nil,-1) ;
for x:=0 to TobEcr.Detail.Count-1 do
  BEGIN
  TOBL:=TOBEcr.Detail[x] ; if TOBL.GetValue('E_ANA')<>'X' then Continue ;
  SQL:='SELECT Y_JOURNAL,Y_EXERCICE,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_NUMLIGNE,Y_AXE,Y_NUMVENTIL,Y_QUALIFPIECE,Y_'+gszZoneECR+' '
      +'FROM ANALYTIQ WHERE Y_JOURNAL="'+TOBL.GetValue('E_JOURNAL')+'" '
      +'AND Y_EXERCICE="'+TOBL.GetValue('E_EXERCICE')+'" AND Y_DATECOMPTABLE="'+UsDateTime(TOBL.GetValue('E_DATECOMPTABLE'))+'" '
      +'AND Y_NUMEROPIECE='+IntToStr(TOBL.GetValue('E_NUMEROPIECE'))+' AND Y_NUMLIGNE='+IntToStr(TOBL.GetValue('E_NUMLIGNE'))+' '
      +'AND Y_QUALIFPIECE="'+TOBL.GetValue('E_QUALIFPIECE')+'"' ;
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then TobAna.LoadDetailDB('ANALYTIQ','','',Q,False,True) ;
  Ferme(Q) ;
  For y:=TobAna.Detail.Count-1 downto 0 do
    TobAna.Detail[y].ChangeParent(TobEcr.Detail[x], -1) ;
  END ;
TobAna.Free ; // FQ 10951
END ;


Procedure TOF_CPAFAUXCPTA.ChargeLignesAModifier ;
var i : integer ;
    St : String ;
Begin
Lecr.Clear ;
Fiche := TFMul(Ecran) ;
  if Fiche.FListe.AllSelected then begin
  Fiche.Q.First;
    While not Fiche.Q.Eof do begin
    St:=Format_String(Fiche.Q.FindField('E_EXERCICE').AsString,3)+';'
       +Format_String(Fiche.Q.FindField('E_JOURNAL').AsString,3)+';'
       +Format_String(Fiche.Q.FindField('E_QUALIFPIECE').AsString,1)+';'
       +DateToStr(Fiche.Q.FindField('E_DATECOMPTABLE').AsDateTime)+';'
       +FormatFloat('0000000000',Fiche.Q.FindField('E_NUMEROPIECE').AsInteger) ;
    LEcr.Add(St) ;
    Fiche.Q.Next ;
    end ;
  end
  else begin
    for i := 0 to Fiche.Fliste.NbSelected -1 do begin
      Fiche.FListe.GotoLeBookmark(i) ;
      St:=Format_String(Fiche.Q.FindField('E_EXERCICE').AsString,3)+';'
       +Format_String(Fiche.Q.FindField('E_JOURNAL').AsString,3)+';'
       +Format_String(Fiche.Q.FindField('E_QUALIFPIECE').AsString,1)+';'
       +DateToStr(Fiche.Q.FindField('E_DATECOMPTABLE').AsDateTime)+';'
       +FormatFloat('0000000000',Fiche.Q.FindField('E_NUMEROPIECE').AsInteger) ;
      LEcr.Add(St) ;
      end ;
    end ;
End ;

 // FQ 10951
procedure TOF_CPAFAUXCPTA.OnArgument (S : String ) ;
begin
THValComboBox(GetControl('E_EXERCICE',True)).OnClick:=ExerciceOnClick ;
TButton(GetControl('BOUVRIR',True)).OnClick:=BOuvrirOnClick ;
SetControlText('E_EXERCICE',VH^.EnCours.Code) ; ExerciceOnClick(nil) ;
THValComboBox(GetControl('E_JOURNAL',True)).ItemIndex := 0;
LEcr:=TStringList.Create ;
LEcr.Sorted:=TRUE ;
LEcr.Duplicates:=dupIgnore ;
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTA.LibereTobs ;
begin
if TobTL<>nil     then begin TobTL.Free;     TobTl:=Nil;     end;
If TobTabLib<>Nil then begin TobTabLib.Free; TobTabLib:=Nil; end;
if TobEcr<>Nil    then begin TobEcr.Free;    TobEcr:=Nil ;   end;
end ;

procedure TOF_CPAFAUXCPTA.OnClose ;
begin
LibereTobs ;
LEcr.Clear ;
LEcr.Free ;
end ;

Initialization
  registerclasses ( [ TOF_CPAFAUXCPTA ] ) ;
end.
