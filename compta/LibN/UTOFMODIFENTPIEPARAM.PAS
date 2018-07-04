{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/10/2002
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
Unit UTOFMODIFENTPIEPARAM ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,  // AGLLanceFiche
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,  // AGLLanceFiche
{$ENDIF}
     sysutils, // DecodeDate,IntToStr
     UTOF,     // TOF
     HCtrls,   // Contrôles Halley
     AGLInit,  // TheTob
     UTob,     // TOB
     HEnt1,    // VK_Valide
     Ent1,     // PositionneEtabUser
     HMsgBox   // PGIBox
     ;

procedure CCLanceFiche_ModifEntPieParam;

Type
  TOF_MODIFENTPIE_PARAM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    RBDate    : TRadioButton;
    RBMois    : TRadioButton;
    gbCanClose : Boolean;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure NewJalOnChange (Sender : TObject) ;
    procedure RbDateOnClick(Sender: TObject);
    procedure RbMoisOnClick(Sender: TObject);
    procedure BOuvrirOnClick(Sender: TObject);
end ;

Implementation

procedure CCLanceFiche_ModifEntPieParam;
begin
  AGLLanceFiche ('CP' , 'CPMODIFENTPIEPARA', '', '' , '' ) ;
end;

procedure TOF_MODIFENTPIE_PARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIE_PARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIE_PARAM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIE_PARAM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIE_PARAM.OnArgument (S : String ) ;
var
  Q : TQuery;
  RSql : String;
  Datedeb,DateFin : TDateTime ;
  wAnnee1,wAnnee2,wMois1,wMois2,wJour1,wJour2 : word ;
  i,iAnnee,iNbAnnee : integer ;
  NewAnnee, NewJal : THValComboBox;
begin
  Inherited ;
  Ecran.HelpContext := 7577200 ;
  gbCanClose:=True;
  Ecran.OnKeyDown := FormKeyDown;
  NewAnnee := THValComboBox(GetControl('NEWANNEE',True));
  NewJal := THValComboBox(GetControl('NEWJAL',True));   If NewJal <> NIL then NewJal.OnChange := NewJalOnChange;
  RBDate := TRadioButton(GetControl('RBDATE',True)); if RbDate <> nil then RbDate.OnClick := RbDateOnClick ;
  RBMois := TRadioButton(GetControl('RBMOIS',True)); if RbMois <> nil then RbMois.OnClick := RbMoisOnClick ;
  TButton(GetControl('BVALIDER',True)).OnClick := BOuvrirOnClick ;

  //SG6 23/11/04 FQ 14227 Suppression des journaux libres et borderaux
  SetControlProperty('NEWJAL','PLUS',LaTOB.GetValue('NEWJALPLUS')+' AND (J_MODESAISIE="-" OR J_MODESAISIE="")');

  //Fin FQ 14227

  SetControlProperty('NEWNATPIE','DATATYPE',LaTOB.GetValue('NATUREDATATYPE'));
  if (LaTOB.GetValue('ETABLISSEMENT')<>'') then THValComboBox(GetControl('NEWETAB',True)).Value := LaTOB.GetValue('ETABLISSEMENT')
                                           else THValComboBox(GetControl('NEWETAB',True)).ItemIndex := 0;
  if (LaTOB.GetValue('JOURNAL')<>'') then THValComboBox(GetControl('NEWJAL',True)).Value := LaTOB.GetValue('JOURNAL')
                                     else THValComboBox(GetControl('NEWJAL',True)).ItemIndex := 0;
  if (LaTOB.GetValue('BLOQNATURE')='X') then SetControlEnabled('NEWNATPIE',False);
  
  if (LaTOB.GetValue('NATURE')<>'') then THValComboBox(GetControl('NEWNATPIE',True)).Value := LaTOB.GetValue('NATURE')
                                    else THValComboBox(GetControl('NEWNATPIE',True)).ItemIndex := 0;
  PositionneEtabUser (THValComboBox(GetControl('NEWETAB',True)), False);
  // Impossibilité de modifier le journal, la nature et la date
  if (LaTOB.GetValue('MODESAISIE')='BOR') then begin
    SetControlEnabled('NEWJAL',False);
    SetControlEnabled('NEWNATPIE',False);
    SetControlEnabled('RBDATE',False);
    SetControlEnabled('NEWDATE',False);
  end;

  // Exercice, mois et annee
  RSql := 'SELECT MIN(EX_DATEDEBUT),MAX(EX_DATEFIN) FROM EXERCICE WHERE EX_ETATCPTA ="OUV"';
  Q :=OpenSql(RSql, True) ;
  DateDeb:=Q.Fields[0].AsDateTime ;
  DateFin:=Q.Fields[1].AsDateTime ;
  Ferme(Q) ;
  Decodedate(DateDeb,wAnnee1,wMois1,wJour1) ;
  DecodeDate(DateFin,wAnnee2,wMois2,wJour2) ; iAnnee:=wAnnee1 ; iNbAnnee:=wAnnee2-wAnnee1 ;
  For i:= 0 to iNbAnnee do BEGIN
    NewAnnee.Values.Add(IntToStr(iAnnee)) ; NewAnnee.Items.add(IntToStr(iAnnee)) ;
    iAnnee:=iAnnee+1 ;
  END ;
//  SetControlProperty('NEWANNEE', 'VALUE', LaTOB.GetValue('ANNEE'));
//  SetControlProperty('NEWMOIS', 'VALUE', LaTOB.GetValue('MOIS'));
  TheTOB := nil;
  
  NewJalOnChange(nil);
end ;

procedure TOF_MODIFENTPIE_PARAM.OnClose ;
begin
  if not gbCanClose then begin gbCanClose:=True; LastError:=-1; end;
  Inherited ;
end ;

procedure TOF_MODIFENTPIE_PARAM.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (csDestroying in Ecran.ComponentState) then Exit;
  if (Key =  vk_valide) then begin Key:=0 ; TButton(GetControl('BVALIDER',True)).Click; end;
end;

procedure TOF_MODIFENTPIE_PARAM.NewJalOnChange(Sender: TObject);
Var Natjal, RSql : String ;
    Q : TQuery ;
BEGIN
  RSql := 'SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+GetControlText('NEWJAL')+'"' ;
  Q :=OpenSql(RSql, True) ;
  NatJal := Q.Fields[0].AsString ;
  Ferme(Q) ;
  SetControlProperty('NEWNATPIE','DATATYPE','TTNATUREPIECE') ;
  If (Natjal= 'BQE') or (NatJal= 'CAI') Then SetControlProperty('NEWNATPIE','DATATYPE','TTNATPIECEBANQUE') else
  If (Natjal= 'ACH')                    Then SetControlProperty('NEWNATPIE','DATATYPE','TTNATPIECEACHAT') else
  If (Natjal= 'VTE')                    Then SetControlProperty('NEWNATPIE','DATATYPE','TTNATPIECEVENTE') else
  If (Natjal= 'OD')                     Then SetControlProperty('NEWNATPIE','DATATYPE','TTJALOD') ;
  THValComboBox(GetControl('NEWNATPIE',True)).ItemIndex := 0 ;
end;

procedure TOF_MODIFENTPIE_PARAM.RbDateOnClick(Sender: TObject);
begin
  SetControlEnabled('NewMois',False);
  SetControlEnabled('NewAnnee',False);
  SetControlEnabled('NewDate',True);
end;

procedure TOF_MODIFENTPIE_PARAM.RbMoisOnClick(Sender: TObject);
begin
  SetControlText('NewDate','01/01/1900');
  SetControlEnabled('NewMois',True);
  SetControlEnabled('NewAnnee',True);
  SetControlEnabled('NewDate',False);
end;

procedure TOF_MODIFENTPIE_PARAM.BOuvrirOnClick(Sender: TObject);
var
  T : TOB;
  dt : TDateTime;
begin
  gbCanClose:=True;
  // Si Mois/Année est coché
  if (RBMois.Checked = True) then begin
    if ((GetControlText('NEWMOIS')='') or (GetControlText('NEWANNEE')='')) then begin
      PGIBox('Vous devez sélectionner un mois et une année.',Ecran.caption);
      gbCanClose := False;
      Exit;
    end;
  end;

  // Si on souhaite modifier la date
  if (RBMois.Checked = True) or (RBDate.Checked = True) then begin
    // On interdit de passer l'écriture sur une période clôturée // 14229
    if (RBMois.Checked = True) then dt := EncodeDate( StrToInt(GetControlText('NEWANNEE')), StrToInt(GetControlText('NEWMOIS')), 1)
                               else dt := StrToDate(GetControlText('NEWDATE'));

    if (dt <= VH^.DateCloturePer) then begin
      PGIBox('Vous ne pouvez pas modifier les écritures en les plaçant sur une période clôturée.',Ecran.caption);
      gbCanClose := False;
      Exit;
    end;
  end;
  
  Inherited ;
  T := TOB.Create ('', nil, -1);
  T.AddChampSupValeur('ETABLISSEMENT',GetControlText('NEWETAB'));
  T.AddChampSupValeur('JOURNAL',GetControlText('NEWJAL'));
  T.AddChampSupValeur('NATURE',GetControlText('NEWNATPIE'));
  T.AddChampSupValeur('MOIS',GetControlText('NEWMOIS'));
  T.AddChampSupValeur('ANNEE',GetControlText('NEWANNEE'));
  T.AddChampSupValeur('DATE',GetControlText('NEWDATE'));
  T.AddChampSupValeur('RBMOIS',RBMois.Checked);
  T.AddChampSupValeur('RBDATE',RBDate.Checked);
  TheTOB := T;
  Ecran.Close;
end;

Initialization
  registerclasses ( [ TOF_MODIFENTPIE_PARAM ] ) ;
end.
