unit TofMajPerSem;

interface

uses Forms,Classes,SysUtils,Controls,StdCtrls,ExtCtrls,filectrl,Hctrls,
     dbtables,HMsgBox,Utof,Utob,ParamSoc,PrintDBG,FE_Main;

procedure CPLanceFiche_MAJPerioSemaine;

type
  TOF_MajPerSem = class(TOF)
  private
    Sortie : Boolean ;
    procedure BStopOnClick(Sender : TObject) ;
    procedure MAjEcriture;
  public
    procedure Onload ; override ;
    procedure OnClose ; override ;
    procedure OnUpdate; override ;
  end;

const TMsg: array[1..4] of string 	= (
      {01}  '1;Mise à jour periode et numéro de semaine;Voulez-vous arrêter le traitement ?;Q;YN;N;N'
      {02} ,'2;Mise à jour periode et numéro de semaine;Traitement en cours. Sortie impossible.;W;O;O;O'
      {03} ,'3;Mise à jour periode et numéro de semaine;Confirmez-vous le recalcul des périodes et semaines ?;Q;YN;N;N'
      {04} ,'4;Mise à jour periode et numéro de semaine;Le traitement s''est correctement terminé.;I;O;O;O'
      );

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  Ent1, HEnt1, HStatus;

procedure CPLanceFiche_MAJPerioSemaine;
begin
  AGLLanceFiche('CP', 'MAJPERIOSEMAIN', '', '', '') ;
end;

procedure TOF_MajPerSem.OnLoad;
var CExercic : THValCOmboBox ; BStop : TButton ;
begin
CExercic:=THValComboBox(GetControl('CEXO'));
BStop:=TButton(GetControl('BSTOP'));
CExercic.ItemIndex:=0 ;
if (BStop<>nil) and (not Assigned(BStop.OnClick)) then BStop.OnClick:=BStopOnClick ;
Sortie:=TRUE ;
end;

procedure TOF_MajPerSem.OnUpdate;
var
  CExercic : THValComboBox ;
  i : Integer;
begin
  SetControlEnabled('PANEL1', False);

  if HShowMessage(Tmsg[3],'','')= mrNo then Exit;

  CExercic:=THValComboBox(GetControl('CEXO', True));
  // Si tous les exercices
  if (CExercic.ItemIndex = 0) then begin
    for i := 1 to CExercic.Items.Count - 1 do begin
      CExercic.ItemIndex := i;
      Application.ProcessMessages;
      MAjEcriture;
    end;
    end
  else begin
    MAjEcriture;
  end;

  SetControlEnabled('PANEL1', True);  
  if not Sortie then HShowMessage(Tmsg[4],'','');
end;

procedure TOF_MajPerSem.BStopOnClick(Sender : TObject);
begin
if Not Sortie then if HShowMessage(TMsg[1],'','')=mrYes then Sortie:=TRUE ;
end;

procedure TOF_MajPerSem.OnClose;
begin
if Not Sortie then
  begin
  HshowMessage(Tmsg[2],'','') ;
  LastError:=1;
  end;
end;

// FQ 10380
procedure TOF_MajPerSem.MAjEcriture;
var
  CPeriode,CSemaine : TCheckBox ;
  CExercic : THValComboBox ;
  i, j, Semaine : integer ;
  DateDeb, DateFin : TDateTime;
  nbPeriode : Integer;
  aa, aa2, mm, mm2 , jj, jj2: Word;
  szSQL, szCodeExo : String;
begin
  Sortie:=FALSE ;
  CExercic:=THValComboBox(GetControl('CEXO', True));
  CPeriode:=TCheckBox(GetControl('CPERIODE', True));
  CSemaine:=TCheckBox(GetControl('CSEMAINE', True));

  DateDeb := 2;
  DateFin := 2;

  if (CPeriode.Checked) or (Csemaine.Checked) then
  begin
    // Récupère les dates début et fin de l'exercice sélectionné
    szCodeExo := CExercic.Values[CExercic.ItemIndex];
    for i := Low(VH^.Exercices) to High(VH^.Exercices) do begin
      if (VH^.Exercices[i].Code = szCodeExo) then begin
        DateDeb := VH^.Exercices[i].Deb;
        DateFin := VH^.Exercices[i].Fin;
        break;
      end;
    end;

    // Traite les périodes
    if (CPeriode.Checked) then begin
      // Calcul le nombre de mois séparant les 2 dates
      DecodeDate(DateDeb, aa, mm ,jj);
      DecodeDate(DateFin, aa2, mm2 ,jj2);
      if (aa= aa2) then
        nbPeriode := (mm2-mm)+1
      else
        nbPeriode := (12-mm) + mm2 +1 + (((aa2-aa)*12)-12);

      // Traitement
      for i := 1 to nbPeriode do begin
        DateFin := FINDEMOIS(DateDeb);
        szSQL := 'UPDATE ECRITURE SET E_PERIODE = '+IntToStr(GetPeriode(DateDeb))+' WHERE E_EXERCICE ="'+szCodeExo+'" AND E_DATECOMPTABLE >="'+USDATETIME(DateDeb)+'" AND E_DATECOMPTABLE <="'+USDATETIME(DateFin)+'"';
        Application.ProcessMessages;
        ExecuteSQL(szSQL);
        szSQL := 'UPDATE ANALYTIQ SET Y_PERIODE = '+IntToStr(GetPeriode(DateDeb))+' WHERE Y_EXERCICE ="'+szCodeExo+'" AND Y_DATECOMPTABLE >="'+USDATETIME(DateDeb)+'" AND Y_DATECOMPTABLE <="'+USDATETIME(DateFin)+'"';
        Application.ProcessMessages;
        ExecuteSQL(szSQL);
        DateDeb := DateFin +1;
      end;
    end;

    // Traite les semaines
    if (CSemaine.Checked) then begin
      // Récupère les dates
      if (CPeriode.Checked) then begin
        DateDeb := EncodeDate(aa, mm ,jj);
        DateFin := EncodeDate(aa2, mm2 ,jj2);
      end;

      // Calcul le nombre de semaine séparant les 2 dates
      nbPeriode := Trunc((DateFin - DateDeb +1) / 7)+1;

      // Traitement
      for i := 1 to nbPeriode do begin
        DateFin := DateDeb +1;
        Semaine := NumSemaine(DateDeb);
        for j := 1 to 7 do begin
          if (NumSemaine(DateFin) = Semaine) then DateFin := DateFin +1
          else begin
            DateFin := DateFin -1;
            break;
          end;
        end;

        szSQL := 'UPDATE ECRITURE SET E_SEMAINE = '+IntToStr(Semaine)+' WHERE E_EXERCICE ="'+szCodeExo+'" AND E_DATECOMPTABLE >="'+USDATETIME(DateDeb)+'" AND E_DATECOMPTABLE <="'+USDATETIME(DateFin)+'"';
        Application.ProcessMessages;
        ExecuteSQL(szSQL);
        szSQL := 'UPDATE ANALYTIQ SET Y_SEMAINE = '+IntToStr(Semaine)+' WHERE Y_EXERCICE ="'+szCodeExo+'" AND Y_DATECOMPTABLE >="'+USDATETIME(DateDeb)+'" AND Y_DATECOMPTABLE <="'+USDATETIME(DateFin)+'"';
        Application.ProcessMessages;
        ExecuteSQL(szSQL);
        DateDeb := DateFin+1
      end;
    end;
  end;
  Sortie:=TRUE ;
end;

Initialization
registerclasses([TOF_MajPerSem]);
end.
