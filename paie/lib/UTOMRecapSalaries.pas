{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 20/06/2001
Modifié le ... : 20/06/2001
Description .. : TOM gestion des cumuls CP et RTT des salariés en
Suite ........ : saisie déportée
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************
PT1 15/07/2002 SB V582 Intégration de compteur validé en attente Cp et Rtt
PT2 04/11/2002 SB V585 Intégration du calcul des compteurs Abs en cours (champs de la table)
}

unit UTOMRecapSalaries;

interface
uses
{$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
   Controls, Classes, sysutils, Hctrls, Utom, Hent1, Utob;
type
  TOM_RecapSalaries = class(TOM)
    procedure OnLoadRecord; override;
    procedure OnArgument(Arguments: string); override;
  private
    theArg: string;
  end;

implementation

uses entpaie; //OPX pgoutilsEagl,


{ TOM_RecapSalaries }

procedure TOM_RecapSalaries.OnArgument(Arguments: string);
begin
  inherited;
  theArg := Arguments;

end;

procedure TOM_RecapSalaries.OnLoadRecord;
var Tob_MvtEncours, TMvt: TOB;
  ATTN, VALIDN, ATT, VALID: Double;
  i: Integer;
  DateClot: TDateTime;
  Q: TQuery;
  St : String;
  LaZone : THEdit;
begin
  inherited;
  LaZone := THEdit (GetControl ('VALIDN'));
  if LaZone <> NIL then
  begin
    SetControlVisible ('VALID', True);
    SetControlVisible ('VALIDN', True);
    SetControlVisible ('ATT', True);
    SetControlVisible ('ATTN', True);
    SetControlVisible ('PRS_PRIVALIDE', False);
    SetControlVisible ('PRS_PRIATTENTE', False);
  end;
  if theArg = 'R' then Ecran.caption := TraduireMemoire('Récapitulatif salarié') + ' ';
  SetControlvisible('PRS_DATEMODIF', False);
  SetControlText('TPRS_DATEMODIF', TraduireMemoire('Date de solde au') + ' ' + DateToStr(VH_Paie.PGECabDateIntegration));
  SetControlEnabled('PRS_CUMRTTREST', False);
  SetControlEnabled('PRS_PRIVALIDE', False);
  SetControlEnabled('PRS_RTTVALIDE', False);
  SetControlEnabled('PRS_PRIATTENTE', False);
  SetControlEnabled('PRS_RTTATTENTE', False);
{PT2 Mise en commentaire : utilisation des champs de la table
If IsNumeric(GetControlText('PRS_CUMRTTACQUIS')) AND IsNumeric(GetControlText('PRS_CUMRTTPRIS')) then
 SetControlText('RESTRTT',FloatToStr(StrToFloat(GetControlText('PRS_CUMRTTACQUIS'))-StrToFloat(GetControlText('PRS_CUMRTTPRIS'))))
else
 SetControlText('RESTRTT','');
//DEB PT1
ValRtt:=0;AttRtt:=0; ValPri:=0; AttPri:=0;
St:='SELECT PCN_TYPECONGE,PCN_VALIDRESP,SUM(PCN_JOURS) JOURS '+
    'FROM ABSENCESALARIE WHERE PCN_EXPORTOK="-" AND PCN_SALARIE="'+GetControlText('PRS_SALARIE')+'" '+
    'GROUP BY PCN_TYPECONGE,PCN_VALIDRESP';
Q :=Opensql(St,True);
While not Q.eof Do
  Begin
  //CP
  if (Q.FindField('PCN_TYPECONGE').AsString='PRI') then
    If (Q.FindField('PCN_VALIDRESP').AsString='VAL') then
        ValPri:=Q.FindField('JOURS').AsFloat
    else
       If (Q.FindField('PCN_VALIDRESP').AsString='ATT') then
          AttPri:=(Q.FindField('JOURS').AsFloat);
  //RTT
  if (Q.FindField('PCN_TYPECONGE').AsString='RTT') then
    If (Q.FindField('PCN_VALIDRESP').AsString='VAL') then
      ValRtt:=Q.FindField('JOURS').AsFloat
    else
       If (Q.FindField('PCN_VALIDRESP').AsString='ATT') then
         AttRtt:=Q.FindField('JOURS').AsFloat;
   //RTS
  if (Q.FindField('PCN_TYPECONGE').AsString='RTS') then
    If (Q.FindField('PCN_VALIDRESP').AsString='VAL') then
      ValRtt:=ValRtt+Q.FindField('JOURS').AsFloat
    else
       If (Q.FindField('PCN_VALIDRESP').AsString='ATT') then
          AttRtt:=AttRtt+Q.FindField('JOURS').AsFloat;
  Q.Next;
  End;
  SetControlText('NBVALIDEPRI',FloatToStr(ValPri));
  SetControlText('NBATTENTEPRI',FloatToStr(AttPri));
  SetControlText('NBVALIDERTT',FloatToSTr(ValRtt));
  SetControlText('NBATTENTERTT',FloatTostr(AttRtt));
//FIN PT1 }
  st := 'SELECT ETB_DATECLOTURECPN FROM ETABCOMPL WHERE ETB_ETABLISSEMENT=(SELECT PSA_ETABLISSEMENT ' +
    'FROM SALARIES WHERE PSA_SALARIE="' + GetControlText('PRS_SALARIE') + '")';
  Q := OpenSql(st, True);
  if not Q.Eof then DateClot := Q.FindField('ETB_DATECLOTURECPN').AsDateTime;
  Ferme(Q);
  st := 'SELECT PCN_JOURS,PCN_DATEFINABS,PCN_VALIDRESP FROM ABSENCESALARIE ' +
    'WHERE (PCN_TYPECONGE="PRI" AND PCN_PERIODECP<2) AND PCN_ETATPOSTPAIE <> "NAN" ' +
    'AND (PCN_EXPORTOK<>"X" OR PCN_CODETAPE="...") AND PCN_SALARIE="' + GetControlText('PRS_SALARIE') + '" ' +
    'ORDER BY PCN_PERIODECP';
  Q := OpenSql(st, True);
  if not Q.Eof then
  begin
      //Charger la tob des mvts en cours
    Tob_MvtEncours := TOb.Create('Les mvts en cours', nil, -1);
    Tob_MvtEncours.LoadDetailDB('Les mvts en cours', '', '', Q, False);
    for i := 0 to Tob_MvtEncours.detail.count - 1 do
    begin
      Tmvt := Tob_MvtEncours.detail[i];
      if (Tmvt.GetValue('PCN_VALIDRESP') = 'VAL') then
      begin
        if (TMvt.GetValue ('PCN_DATEFINABS') <= DateClot ) then VALIDN := VALIDN + Tmvt.GetValue('PCN_JOURS')
        else VALID := VALID + Tmvt.GetValue('PCN_JOURS');
      end
      else
        if (Tmvt.GetValue('PCN_VALIDRESP') = 'ATT') then
        begin
        if (TMvt.GetValue ('PCN_DATEFINABS') <= DateClot ) then ATTN := ATTN + Tmvt.GetValue('PCN_JOURS')
        else ATT := ATT + Tmvt.GetValue('PCN_JOURS');
        end;
    end;
  end;
  Ferme(Q);
  FreeAndNil (Tob_MvtEncours);
  SetControlText ('ATT' , StrfMontant(ATT, 15, 2, '', TRUE));
  SetControlText ('ATTN' , StrfMontant(ATTN, 15, 2, '', TRUE));
  SetControlText ('VALID' , StrfMontant(VALID, 15, 2, '', TRUE));
  SetControlText ('VALIDN' , StrfMontant(VALIDN, 15, 2, '', TRUE));
end;

initialization
  registerclasses([TOM_RecapSalaries]);
end.

