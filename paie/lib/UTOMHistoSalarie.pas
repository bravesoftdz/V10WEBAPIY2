{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion historique des évènements salariés
Mots clefs ... : PAIE;SALARIE
*****************************************************************}
{
PT1 : 25/07/01 : V540 : VG Ajout champs dans Historique salarié
PT2 : 15/10/01 : V562 : VG Ajout champs dans Historique salarié
PT3 : 07/02/02 : V571 : PH suppression fonctions historisation
PT4 : 25/03/02 : V571 : PH modifs champs
PT5 : 22/10/02 : V585 : PH Rend les champs invisibles si non gérés
PT6 21/12/2007 V_81 FC Concept accessibilité fiche salarié
}
unit UTOMHistoSalarie;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fiche,
{$ELSE}
      eFiche,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,
      ParamSoc,HPanel,ParamDat,Dialogs,PgCongesPayes,PgOutils,
      AGLInit,P5util,EntPaie;

Type

     TOM_HistoSalarie = Class (TOM)
       procedure OnChangeField (F : TField); override ;
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnUpdateRecord ;                    override ;
// PT3       procedure AlimenteHistoSalTob;
       procedure OnNewRecord  ;                      override ;
       procedure OnLoadRecord ; override ;
     private
       salarie: string;
//       Action : string;
//       Modif  : boolean;
       function IncNoOrdrePHS(sal: string; DD, DF : TDateTime): Integer;
       procedure BValiderOnClick(Sender: TObject);
       procedure AfficheTitre;
       procedure enabledchampOngletHistorique;
       procedure ModifClick(Sender : tObject);
     END ;


implementation


procedure TOM_HistoSalarie.OnArgument(stArgument: String);
var
   Arg     : string;
   TModif  : ttoolbarbutton97;
   Action:String; //PT6
begin
Inherited ;
 Arg := stArgument;
 Salarie := Arg;
 salarie:=Trim(ReadTokenPipe(Arg,';')) ;
 Action := ''; //PT6
 if Arg <> '' then
   Action:=Trim(ReadTokenPipe(Arg,';')) ; //PT6
// Action:=Trim(ReadTokenPipe(Arg,';')) ;
{ Modif := false;

 if Action = 'CREATION' then
      begin
      TFFiche(Ecran).BinsertClick(nil);
      Modif := true;
      end;
}
 AfficheTitre;
 enabledchampOngletHistorique;
{ TModif  :=ttoolbarbutton97(getcontrol('TMODIF'));
 if TModif <> nil then
    Tmodif.onclick := ModifClick;
 }

  //DEB PT6
  if Action='CONSULTATION' then
  begin
    SetControlProperty('PGeneral','enabled',false);
    SetControlProperty('SALAIRES','enabled',false);
    SetControlProperty('ZLIBRE','enabled',false);
    SetControlProperty('DADS','enabled',false);
    SetControlProperty('BInsert','enabled',false);
    SetControlProperty('BDelete','enabled',false);
    SetControlProperty('BValider','enabled',false);
  end;
  //FIN PT6
end;



procedure TOM_HistoSalarie.ModifClick(Sender : tObject);
begin
// Modif := true;
 enabledchampOngletHistorique;
end;
//PT3 07/02/2002 V571 PH suppression historisation salarié pendant la saisie salarié
{procedure TOM_HistoSalarie.AlimenteHistoSalTob;
begin
 SetField('PHS_CODEEMPLOI'     , TheHistoSal.CodeEmploi);
 SetField('PHS_LIBELLEEMPLOI'  , TheHistoSal.LibelleEmploi);
 SetField('PHS_QUALIFICATION'  , TheHistoSal.Qualification);
 SetField('PHS_COEFFICIENT'    , TheHistoSal.Coefficient);
 SetField('PHS_INDICE'         , TheHistoSal.Indice);
 SetField('PHS_NIVEAU'         , TheHistoSal.Niveau);
 SetField('PHS_CODESTAT'       , TheHistoSal.CodeStat);
 SetField('PHS_TRAVAILN1'      , TheHistoSal.TravailN1);
 SetField('PHS_TRAVAILN2'      , TheHistoSal.TravailN2);
 SetField('PHS_TRAVAILN3'      , TheHistoSal.TravailN3);
 SetField('PHS_TRAVAILN4'      , TheHistoSal.TravailN4);
 SetField('PHS_GROUPEPAIE'     , TheHistoSal.Groupepaie);
 SetField('PHS_SALAIREMOIS1'   , TheHistoSal.SalaireMois1);
 SetField('PHS_SALAIREANN1'    , TheHistoSal.SalaireAnn1);
 SetField('PHS_SALAIREMOIS2'   , TheHistoSal.SalaireMois2);
 SetField('PHS_SALAIREANN2'    , TheHistoSal.SalaireAnn2);
 SetField('PHS_SALAIREMOIS3'   , TheHistoSal.SalaireMois3);
 SetField('PHS_SALAIREANN3'    , TheHistoSal.SalaireAnn3);
 SetField('PHS_SALAIREMOIS4'   , TheHistoSal.SalaireMois4);
 SetField('PHS_SALAIREANN4'    , TheHistoSal.SalaireAnn4);
 SetField('PHS_DTLIBRE1'       , TheHistoSal.Dtlibre1);
 SetField('PHS_DTLIBRE2'       , TheHistoSal.Dtlibre2);
 SetField('PHS_DTLIBRE3'       , TheHistoSal.Dtlibre3);
 SetField('PHS_DTLIBRE4'       , TheHistoSal.Dtlibre4);
 SetField('PHS_BOOLLIBRE1'     , TheHistoSal.Boollibre1);
 SetField('PHS_BOOLLIBRE2'     , TheHistoSal.Boollibre2);
 SetField('PHS_BOOLLIBRE3'     , TheHistoSal.Boollibre3);
 SetField('PHS_BOOLLIBRE4'     , TheHistoSal.Boollibre4);
 SetField('PHS_CBLIBRE1'       , TheHistoSal.CSlibre1);
 SetField('PHS_CBLIBRE2'       , TheHistoSal.CSlibre2);
 SetField('PHS_CBLIBRE3'       , TheHistoSal.CSlibre3);
 SetField('PHS_CBLIBRE4'       , TheHistoSal.CSlibre4);
 SetField('PHS_PROFIL'         , TheHistoSal.profil);
 SetField('PHS_PERIODBUL'      , TheHistoSal.periodbul);
 SetField('PHS_PGHHORAIREMOIS' , TheHistoSal.Horairemois);
 SetField('PHS_PGHHORHEBDO'    , TheHistoSal.Horairehebdo);
 SetField('PHS_PGHHORANNUEL'   , TheHistoSal.HoraireAnnuel);
 SetField('PHS_PGHTAUXHORAIRE' , TheHistoSal.TauxHoraire);
//PT1
 SetField('PHS_DADSPROF'       , TheHistoSal.DADSProfessio);
 SetField('PHS_DADSCAT'        , TheHistoSal.DADSCategorie);
 SetField('PHS_TAUXPARTIEL'    , TheHistoSal.TauxPartiel);
 SetField('PHS_CONDEMPLOI'     , TheHistoSal.Condemploi);//PT2

 SetField('PHS_BCODEEMPLOI'     , TheHistoSal.BCodeEmploi);
 SetField('PHS_BLIBELLEEMPLOI'  , TheHistoSal.BLibelleEmploi);
 SetField('PHS_BQUALIFICATION'  , TheHistoSal.BQualification);
 SetField('PHS_BCOEFFICIENT'    , TheHistoSal.BCoefficient);
 SetField('PHS_BINDICE'         , TheHistoSal.BIndice);
 SetField('PHS_BNIVEAU'         , TheHistoSal.BNiveau);
 SetField('PHS_BCODESTAT'       , TheHistoSal.BCodeStat);
 SetField('PHS_BTRAVAILN1'      , TheHistoSal.BTravailN1);
 SetField('PHS_BTRAVAILN2'      , TheHistoSal.BTravailN2);
 SetField('PHS_BTRAVAILN3'      , TheHistoSal.BTravailN3);
 SetField('PHS_BTRAVAILN4'      , TheHistoSal.BTravailN4);
 SetField('PHS_BGROUPEPAIE'     , TheHistoSal.BGroupepaie);
 SetField('PHS_BSALAIREMOIS1'   , TheHistoSal.BSalaireMois1);
 SetField('PHS_BSALAIREANN1'    , TheHistoSal.BSalaireAnn1);
 SetField('PHS_BSALAIREMOIS2'   , TheHistoSal.BSalaireMois2);
 SetField('PHS_BSALAIREANN2'    , TheHistoSal.BSalaireAnn2);
 SetField('PHS_BSALAIREMOIS3'   , TheHistoSal.BSalaireMois3);
 SetField('PHS_BSALAIREANN3'    , TheHistoSal.BSalaireAnn3);
 SetField('PHS_BSALAIREMOIS4'   , TheHistoSal.BSalaireMois4);
 SetField('PHS_BSALAIREANN4'    , TheHistoSal.BSalaireAnn4);
 SetField('PHS_BDTLIBRE1'       , TheHistoSal.BDtlibre1);
 SetField('PHS_BDTLIBRE2'       , TheHistoSal.BDtlibre2);
 SetField('PHS_BDTLIBRE3'       , TheHistoSal.BDtlibre3);
 SetField('PHS_BDTLIBRE4'       , TheHistoSal.BDtlibre4);
 SetField('PHS_BBOOLLIBRE1'     , TheHistoSal.BBoollibre1);
 SetField('PHS_BBOOLLIBRE2'     , TheHistoSal.BBoollibre2);
 SetField('PHS_BBOOLLIBRE3'     , TheHistoSal.BBoollibre3);
 SetField('PHS_BBOOLLIBRE4'     , TheHistoSal.BBoollibre4);
 SetField('PHS_BCBLIBRE1'       , TheHistoSal.BCSlibre1);
 SetField('PHS_BCBLIBRE2'       , TheHistoSal.BCSlibre2);
 SetField('PHS_BCBLIBRE3'       , TheHistoSal.BCSlibre3);
 SetField('PHS_BCBLIBRE4'       , TheHistoSal.BCSlibre4);
 SetField('PHS_BPROFIL'         , TheHistoSal.BProfil);
 SetField('PHS_BPERIODBUL'      , TheHistoSal.BPeriodBul);
 SetField('PHS_BPGHHORMOIS'     , TheHistoSal.BHorairemois);
 SetField('PHS_BPGHHORHEBDO'    , TheHistoSal.BHorairehebdo);
 SetField('PHS_BPGHHORANNUEL'   , TheHistoSal.BHoraireAnnuel);
 SetField('PHS_BPGHTAUXHOR'     , TheHistoSal.BTauxHoraire);
//PT1
 SetField('PHS_BDADSPROF'       , TheHistoSal.BDADSProfessio);
 SetField('PHS_BDADSCAT'        , TheHistoSal.BDADSCategorie);
 SetField('PHS_BTAUXPARTIEL'    , TheHistoSal.BTauxPartiel);
 SetField('PHS_BCONDEMPLOI'     , TheHistoSal.BCondemploi);//PT2
end;
}

procedure TOM_HistoSalarie.OnNewRecord;
begin
inherited;
// Action := 'CREATION';
// TFFiche(Ecran).FTypeAction := taCreat ;
// PT4 : 25/03/02 : V571 : PH modifs champs
 SetField('PHS_DATEAPPLIC'     ,Date );
 SetField('PHS_DATEEVENEMENT'       ,Date );
 SetField('PHS_DATERETRO'       ,Date );
 SetField('PHS_BCODEEMPLOI', '-');
 SetField('PHS_BLIBELLEEMPLOI', '-');
 SetField('PHS_BQUALIFICATION', '-');
 SetField('PHS_BCOEFFICIENT', '-');
 SetField('PHS_BINDICE', '-');
 SetField('PHS_BNIVEAU', '-');
 SetField('PHS_BCODESTAT', '-');
 SetField('PHS_BTRAVAILN1', '-');
 SetField('PHS_BTRAVAILN2', '-');
 SetField('PHS_BTRAVAILN3', '-');
 SetField('PHS_BTRAVAILN4', '-');
 SetField('PHS_BSALAIREMOIS1', '-');
 SetField('PHS_BSALAIREMOIS2', '-');
 SetField('PHS_BSALAIREMOIS3', '-');
 SetField('PHS_BSALAIREMOIS4', '-');
 SetField('PHS_BSALAIREMOIS5', '-');
 SetField('PHS_BDTLIBRE1', '-');
 SetField('PHS_BDTLIBRE2', '-');
 SetField('PHS_BDTLIBRE3', '-');
 SetField('PHS_BDTLIBRE4', '-');
 SetField('PHS_BBOOLLIBRE1', '-');
 SetField('PHS_BBOOLLIBRE2', '-');
 SetField('PHS_BBOOLLIBRE3', '-');
 SetField('PHS_BBOOLLIBRE4', '-');
 SetField('PHS_BCBLIBRE1', '-');
 SetField('PHS_BCBLIBRE2', '-');
 SetField('PHS_BCBLIBRE3', '-');
 SetField('PHS_BCBLIBRE4', '-');
 SetField('PHS_BPROFIL', '-');
 SetField('PHS_BPGHHORHEBDO', '-');
 SetField('PHS_BPGHHORANNUEL', '-');
 SetField('PHS_BDADSPROF', '-');
 SetField('PHS_BDADSCAT', '-');
 SetField('PHS_BPROFILREM', '-');
 SetField('PHS_BCONDEMPLOI', '-');
 SetField('PHS_BHORAIREMOIS', '-');
 SetField('PHS_BTAUXHORAIRE', '-');
 SetField('PHS_BSALAIRETHEO', '-');
 SetField('PHS_BPERIODBUL', '-');
 SetField('PHS_BTTAUXPARTIEL', '-');
 SetField('PHS_BCHARLIBRE1', '-');
 SetField('PHS_BCHARLIBRE2', '-');
 SetField('PHS_BCHARLIBRE3', '-');
 SetField('PHS_BCHARLIBRE4', '-');
end;

procedure TOM_HistoSalarie.BValiderOnClick(Sender: TObject);
begin
end;
procedure TOM_HistoSalarie.OnChangeField (F : TField);
begin
Inherited ;
{
if F.FieldName = 'PHS_SALARIE' then
 Affichetitre;}
end;
procedure TOM_HistoSalarie.AfficheTitre;
var
q: tquery;
nom,prenom : string;
st : string;
begin
Salarie := GetControlText ('PHS_SALARIE');
      q:= opensql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE ="'+Salarie+'"',True);
      if not q.eof then
         begin
         Nom := q.findfield('PSA_LIBELLE').AsString;
         Prenom := q.findfield('PSA_PRENOM').AsString;
         TFFiche(Ecran).caption := 'Salarié : '+salarie+' '+Nom+' '+Prenom+' paie du : '+
           getcontrolText ('PHS_DATEEVENEMENT')+' au '+getcontrolText ('PHS_DATEAPPLIC');
         UpdateCaption(TFFiche(Ecran));
         end;
      Ferme(Q);
end;
procedure TOM_HistoSalarie.OnUpdateRecord;
var i : Integer;
begin
Inherited ;

if  TFFiche(Ecran).FTypeAction = taCreat then
    begin
    i := IncNoOrdrePHS (GetField ('PHS_SALARIE'),GetField ('PHS_DATEAPPLIC'),getfield('PHS_DATEEVENEMENT'));
    SetField ('PHS_ORDRE',i);
    end;
// PT4 : 25/03/02 : V571 : PH modifs champs
{
if getfield('PHS_DATEAPPLIC') > getfield('PHS_DATEEVENEMENT') then
   begin
     LastError:=1;
     LastErrorMsg:='Attention, cette date doit être antérieure ou égale à la date du jour';
   end;
}
end;
procedure TOM_HistoSalarie.enabledchampOngletHistorique;
var
num ,i              : integer;
Numero,libelle      : string;
TLieum,TLieua,TLieu : THLabel;
begin
{
if DS <> NIL then
   begin
   for i := 0 to DS.FieldCount-1 do
    begin
    if pos('PHS_',DS.Fields[i].FieldName) > 0 then
        setcontrolEnabled(ds.Fields[i].FieldName,modif);
    end;
   end;}
// PT5 : 22/10/02 : V585 : PH Rend les champs invisibles si non gérés
// et controle de tous les noms des champs qui étaient faux
For num := 1 to 4 do
  begin
  Numero:=InttoStr(num);
  SetControlVisible('TPHS_TRAVAILN'+Numero,FALSE);
  SetControlVisible('PHS_TRAVAILN'+Numero,FALSE);
  end;
For num := 1 to 5 do
  begin
  Numero:=InttoStr(num);
  SetControlVisible('TPHS_SALAIREMOIS'+Numero,FALSE);
  SetControlVisible('TPHS_SALAIREANN'+Numero,FALSE);
  SetControlVisible('PHS_SALAIREMOIS'+Numero,FALSE);
  SetControlVisible('PHS_SALAIREANN'+Numero,FALSE);
  end;
For num := 1 to 4 do
  begin
  Numero:=InttoStr(num);
  SetControlVisible('TPHS_DTLIBRE'+Numero, FALSE);
  SetControlVisible('PHS_DTLIBRE'+Numero, FALSE);
  end;
For num := 1 to 4 do
  begin
  Numero:=InttoStr(num);
  SetcontrolVisible('PHS_BOOLLIBRE'+Numero,FALSE);
  end;
For num := 1 to 4 do
  begin
  Numero:=InttoStr(num);
  SetcontrolVisible('TPHS_CBLIBRE'+Numero, FALSE);
  SetcontrolVisible('PHS_CBLIBRE'+Numero, FALSE);
  end;
For num := 1 to VH_Paie.PGNbreStatOrg do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PGNbreStatOrg then break;
  libelle:='';
  if Num = 1 then libelle :=VH_Paie.PGLibelleOrgStat1;
  if Num = 2 then libelle :=VH_Paie.PGLibelleOrgStat2;
  if Num = 3 then libelle :=VH_Paie.PGLibelleOrgStat3;
  if Num = 4 then libelle :=VH_Paie.PGLibelleOrgStat4;
  if (Libelle<>'') then
    Begin
    TLieu:=THLabel(GetControl('TPHS_TRAVAILN'+Numero));
    if (TLieu <> NIL) then
      begin
      setcontrolVisible('TPHS_TRAVAILN'+Numero, true);
      TLieu.caption := libelle;
      setcontrolVisible('PHS_TRAVAILN'+Numero, true);
      end;
    End; // begin libelle
  end;

For num := 1 to VH_Paie.PgNbSalLib do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbSalLib then break;
  TLieum:=THLabel(GetControl('TPHS_SALAIREMOIS'+Numero));
  TLieua:=THLabel(GetControl('TPHS_SALAIREANN'+Numero));
  if ((TLieum <> NIL) and (TLieua <> NIL))  then
   begin
   setcontrolVisible('PHS_SALAIREMOIS'+Numero, true);
   setcontrolVisible('PHS_SALAIREANN'+Numero, true);
   TLieua.Visible :=TRUE;
   TLieum.Visible :=TRUE;
   if Num = 1 then
      Begin
      TLieum.Caption :=VH_Paie.PgSalLib1 +' (mensuel)';
      TLieua.Caption :=VH_Paie.PgSalLib1 +' (annuel)';
      end;
   if Num = 2 then
      Begin
      TLieum.Caption :=VH_Paie.PgSalLib2 +' (mensuel)';
      TLieua.Caption :=VH_Paie.PgSalLib2 +' (annuel)';
      end;
   if Num = 3 then
      Begin
      TLieum.Caption :=VH_Paie.PgSalLib3 +' (mensuel)';
      TLieua.Caption :=VH_Paie.PgSalLib3 +' (annuel)';
      end;
   if Num = 4 then
      Begin
      TLieum.Caption :=VH_Paie.PgSalLib4 +' (mensuel)';
      TLieua.Caption :=VH_Paie.PgSalLib4 +' (annuel)';
      end;
   if Num = 5 then
      Begin
      TLieum.Caption :=VH_Paie.PgSalLib5 +' (mensuel)';
      TLieua.Caption :=VH_Paie.PgSalLib5 +' (annuel)';
      end;
   end;
  end;

For num := 1 to VH_Paie.PgNbDate do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbDate then break;
  TLieu:=THLabel(GetControl('TPHS_DTLIBRE'+Numero));
  if (Tlieu <> NIL)   then
   begin
   TLieu.Visible :=TRUE;
   setcontrolVisible('PHS_DTLIBRE'+Numero, true);
   if Num = 1 then Begin TLieu.Caption :=VH_Paie.PgLibDate1; end;
   if Num = 2 then Begin TLieu.Caption :=VH_Paie.PgLibDate2; end;
   if Num = 3 then Begin TLieu.Caption :=VH_Paie.PgLibDate3; end;
   if Num = 4 then Begin TLieu.Caption :=VH_Paie.PgLibDate4; end;
   end;
  end;    //Fin For

For num := 1 to VH_Paie.PgNbCoche do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbCoche then break;
  TLieu:=THLabel(GetControl('PHS_BOOLLIBRE'+Numero));
  SetcontrolVisible('PHS_BOOLLIBRE'+Numero,true);
  TLieu.Visible :=TRUE;
  if Num = 1 then Begin TLieu.Caption :=VH_Paie.PgLibCoche1; end;
  if Num = 2 then Begin TLieu.Caption :=VH_Paie.PgLibCoche2; end;
  if Num = 3 then Begin TLieu.Caption :=VH_Paie.PgLibCoche3; end;
  if Num = 4 then Begin TLieu.Caption :=VH_Paie.PgLibCoche4; end;
  end;

For num := 1 to VH_Paie.PgNbCombo do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbCombo then break;
  TLieu := THLabel(GetControl('TPHS_CBLIBRE'+Numero));
  if (TLieu <> NIL)   then
   begin
   SetcontrolVisible('PHS_CBLIBRE'+ numero,true);
   TLieu.Visible :=TRUE;
   if Num = 1 then Begin TLieu.Caption :=VH_Paie.PgLibCombo1; end;
   if Num = 2 then Begin TLieu.Caption :=VH_Paie.PgLibCombo2; end;
   if Num = 3 then Begin TLieu.Caption :=VH_Paie.PgLibCombo3; end;
   if Num = 4 then Begin TLieu.Caption :=VH_Paie.PgLibCombo4; end;
   end;
  end;
// FIN PT5
end;

function TOM_HistoSalarie.IncNoOrdrePHS(sal: string; DD, DF : TDateTime): Integer;
var
   q: TQuery;
begin
    result := 1;
    q := OpenSQL('select max (PHS_ORDRE) as maxno from HISTOSALARIE ' +
                 ' where PHS_SALARIE = "' + sal+'" AND PHS_DATEEVENEMENT="'+UsDateTime(DD)+'" AND PHS_DATEAPPLIC="'+
                 UsDateTime(DF)+'"', TRUE);

    if not q.eof then result := Q.FindField ('maxno').AsInteger;
    Ferme(q);
end;

procedure TOM_HistoSalarie.OnLoadRecord;
begin
  inherited;
AfficheTitre;
end;

Initialization
registerclasses([TOM_HistoSalarie]) ;

end.
