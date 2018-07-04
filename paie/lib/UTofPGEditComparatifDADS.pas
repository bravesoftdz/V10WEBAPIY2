{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGEDIT_DADSPER ()
Mots clefs ... : TOF;PGEDIT_DADSPER
*****************************************************************}
{
PT1   : 10/08/2004 VG V_50 Edition BTP - FQ N°11317
PT2   : 15/09/2004 VG V_50 Ne pas afficher les écarts <=1 (écart d'arrondi)
                           FQ N°11317
PT3   : 01/02/2005 VG V_60 Edition en CWAS
PT4   : 02/03/2005 VG V_60 LanceEtatTob propriété de QRS1 - FQ N°11955
PT5   : 25/11/2005 JL V_65 FQ 12716 Suppression du message aucune différence a éditer car géré par AGL, et lançait l'état sans tob.
PT6   : 09/01/2006 PH V_65 FQ 12281 Multi-période DADS_U
PT7-1 : 17/10/2006 VG V_70 Remplacement des ellipsis et combo par des multi-val
                           combo
PT7-2 : 17/10/2006 VG V_70 Gestion du décalage de paie - FQ N°12860
PT8   : 30/01/2007 V_80 FC Mise en place filtrage des habilitations/poupulations
                           correction le 27/03/07
PT9   : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT10  : 03/12/2007 VG V_80 Prise en compte du trimestre civil dans le cas d'une
                           déclaration trimestrielle - FQ N°13245
}
unit UTofPGEditComparatifDADS;

interface
uses StdCtrls,
     Controls,
     Classes,
     sysutils,
     ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,
     UtileAgl,
     MaineAgl,
{$ELSE}
     FE_Main,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     QRS1,
{$ENDIF}
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     ParamSoc,
     UTOB,
     PGoutils2,
     ed_tools,
     HStatus,
     HTB97,
     PgDADSCommun,
     P5Def //MonHabilitation
     ;


type
  TOF_PGEDITCOMPDADS = class(TOF)
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;

    private
    TobEtat : TOB;
    Annee, Car : THValComboBox;
    procedure CocheSansDetail(Sender : TObject);
    procedure DateDecaleChange(Sender: TObject);
    procedure Parametrage(Sender: TObject);

  end;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITCOMPDADS.OnArgument;
var
AnneeE, AnneePrec, ComboExer, MoisE, StPlus: string;
AnneeA, Jour, MoisM: Word;
JourJ: TDateTime;
Check : TCheckBox;
ModifDateDecale : TToolbarButton97;
begin
inherited;
{PT7-1
RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
SetControlText('ETAB1', Min);
SetControlText('ETAB2', Max);
RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
SetControlText('Sal1', Min);
SetControlText('Sal2', Max);
}
JourJ:= Date;
DecodeDate(JourJ, AnneeA, MoisM, Jour);
if MoisM > 9 then
   AnneePrec:= IntToStr(AnneeA)
else
   AnneePrec:= IntToStr(AnneeA - 1);
{PT7-1
if RendExerSocialPrec(MoisE, AnneeE, ComboExer, Deb, Fin, AnneePrec) = TRUE then
}
if RendExerSocialPrec(MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec) = TRUE then
   SetControlText('ANNEE', ComboExer)
else
   PGIBox('L''exercice '+AnneePrec+' n''existe pas', Ecran.Caption);

{PT7-1
SetControlText('L_DDU', DateToStr(Deb));
SetControlText('L_DAU', DateToStr(Fin));
SetControlText('DATEDEB', DateToStr(Deb));
SetControlText('DATEFIN', DateToStr(Fin));
}
SetControlText('DATEDEB', DateToStr(DebExer));
SetControlText('DATEFIN', DateToStr(FinExer));

{PT7-1
Edit := THEdit(GetControl('SAL1'));
If Edit <> Nil then Edit.OnExit := ExitEdit;
Edit := THEdit(GetControl('SAL2'));
If Edit <> Nil then Edit.OnExit := ExitEdit;
}
SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
Check := TCheckBox(GetControl('CDETAILSAL'));
If Check <> Nil then
   Check.OnClick := CocheSansDetail;

//PT7-2
// Gestion du bouton de modification des dates
ModifDateDecale:= TToolbarButton97(GetControl('BMODIFDATE'));
if ModifDateDecale<>nil then
   ModifDateDecale.OnClick:= DateDecaleChange;

Check:= TCheckBox (GetControl ('CBTP'));
If Check <> NIL then
   Check.OnClick := Parametrage;

Annee:= THValComboBox (GetControl ('ANNEE'));
If Annee <> NIL then
   Annee.OnChange := Parametrage;

Car:= THValComboBox(GetControl('CCAR'));
if (Car <> nil) then
   begin
   Car.OnChange:= Parametrage;
   Car.Value:= 'A00';
   end;
//FIN PT7-2

//PT9
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('ETAB', 'Plus', StPlus);
//FIN PT9
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITCOMPDADS.OnUpdate;
var
Q: TQuery;
TobDADS, TobEtabliss, T, TobCum, TED: Tob;
MtCum, MtDADS: Double;
TabDads, TabCumul: array[1..5] of string;
i, j, x, a, Count : Integer;
Etablissement, Salarie, StEtab, StPages, StSal, StSalDADS : string;
StSalFI, StSQL, EtabSelect, MCEtab, MCSal : string;
Pages: TPageControl;
Difference, RuptEtablissement : Boolean;
Habilitation:String;//PT8
LesEtablis,Etabl : String; //PT8
k,l : integer;              //PT8
begin
inherited;
{PT7-1
Q := OpenSQL('SELECT PEX_DATEDEBUT,PEX_DATEFIN, PEX_ANNEEREFER'+
             ' FROM EXERSOCIAL WHERE'+
             ' PEX_EXERCICE="' + GetControlText('ANNEE') + '"', True);
if not Q.eof then
   begin
   if (GetControlText('CBTP')='X') then
      begin
      AnneeRefer:= Q.FindField('PEX_ANNEEREFER').AsString;
      DebExer := StrToDate('01/04/'+IntToStr(StrToInt(AnneeRefer)-1));
      FinExer := StrToDate('31/03/'+IntToStr(StrToInt(AnneeRefer)));
      end
   else
      begin
      DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
      FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
      if (DateDeb <> StrToDate (GetControltext ('DATEDEB'))) then DateDeb := StrToDate (GetControltext ('DATEDEB'));
      if (DateFin <> StrToDate (GetControltext ('DATEFIN'))) then DateFin := StrToDate (GetControltext ('DATEFIN'));
      end;
   end
else
   begin
   DebExer:=IDate1900;
   FinExer:=IDate1900;
   end;
Ferme(Q);
SetControlText('DATEDEB',DateToStr(DebExer));
SetControlText('DATEFIN',DateToStr(FinExer));
SetControlText('L_DDU',DateToStr(DateDeb));
SetControlText('L_DAU',DateToStr(DateFin));
}

Pages := TPageControl(GetControl('PAGES'));

// Nombre d'heures travaillées
TabDads[1] := 'S41.G01.00.021';
tabCumul[1] := '20';

// Nombre d'heures salariées
TabDads[2] := 'S41.G01.00.022';
tabCumul[2] := '21';

// Net à payer
TabDads[3] := 'S41.G01.00.063.001';
tabCumul[3] := '09';

// Base brut fiscale
TabDads[4] := 'S41.G01.00.035.001';
tabCumul[4] := '02';

// Base CSG
TabDads[5] := 'S41.G01.00.032.001';
tabCumul[5] := '15';

if GetCheckBoxState('CRUPTETAB') = CbChecked then
   RuptEtablissement := True
else
   RuptEtablissement := False;

FreeAndNil (TobEtat);

if RuptEtablissement = true then
   begin
//PT7-1
   if (THMultiValCombobox (GetControl ('ETAB')).Tous) then
      begin
      StEtab:= ' WHERE ET_FICTIF<>"X"';                //PT9
//DEB PT8
      if (Assigned (MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
         begin
         LesEtablis:= MonHabilitation.LesEtab;
         Etabl:= ReadTokenSt (LesEtablis);
         Habilitation:= '';
         k:= 0;
         while (Etabl<>'') do
               begin
               k:= k+1;
               if (Etabl<>'') then
                  begin
                  if (k>1) then
                     Habilitation:= Habilitation+',';
                  Habilitation:= Habilitation+'"'+Etabl+'"';
                  end;
               Etabl:= ReadTokenSt (LesEtablis);
               end;
         if (k>0) then
            StEtab:= StEtab+' AND'+
                     ' ET_ETABLISSEMENT IN ('+Habilitation+MonHabilitation.SqlPop+')';
         end;
//FIN PT8
      end
   else
      begin
      MCEtab:= GetControlText('ETAB');
      EtabSelect:= ReadTokenpipe(MCEtab, ';');
      StEtab:= ' WHERE (';
      While (EtabSelect <> '') do
            begin
            StEtab:= StEtab+' ET_ETABLISSEMENT="'+EtabSelect+'"';
            EtabSelect := ReadTokenpipe(MCEtab,';');
            if (EtabSelect <> '') then
               StEtab:= StEtab+' OR'
            else
               StEtab:= StEtab+')';
            end;
      end;
//FIN PT7-1
   Q := OpenSQL('SELECT ET_ETABLISSEMENT,ET_LIBELLE'+
                ' FROM ETABLISS'+StEtab, True);

   TobEtabliss := Tob.Create('MesEtablissements', nil, -1);
   TobEtabliss.LoadDetailDB('MesEtablissements', '', '', Q, False);
   Ferme(Q);
   TobEtat := Tob.Create('Edition', nil, -1);
   InitMoveProgressForm (nil, 'Chargement des données',
                         'Veuillez patienter SVP ...',
                         TobEtabliss.Detail.Count, False, True);
   InitMove(TobEtabliss.Detail.Count, '');
{PT7-1
   if GetControlText('SAL1')<>'' then
      StSalDADS:= ' AND D1.PDS_SALARIE>="'+GetControlText('SAL1')+'" AND'+
                  ' D1.PDS_SALARIE<="' + GetControlText('SAL2') + '"'
   else
      StSalDADS:= '';

   if GetControlText('SAL1')<>'' then
      StSalFI:= ' AND PHC_SALARIE>="'+GetControlText('SAL1')+'" AND'+
                ' PHC_SALARIE<="'+GetControlText('SAL2')+'"'
   else
      StSalFI := '';
}
   if (THMultiValCombobox(GetControl('SAL')).Tous) then
      StSalDADS:= ''
   else
      begin
      MCSal:= GetControlText('SAL');
      Salarie:= ReadTokenpipe(MCSal, ';');
      StSalDADS:= ' AND (';
      StSalFI:= ' AND (';
      While (Salarie <> '') do
            begin
            StSalDADS:= StSalDADS+' D1.PDS_SALARIE="'+Salarie+'"';
            StSalFI:= StSalFI+' PHC_SALARIE="'+Salarie+'"';
            Salarie := ReadTokenpipe(MCSal,';');
            if (Salarie <> '') then
               begin
               StSalDADS:= StSalDADS+' OR';
               StSalFI:= StSalFI+' OR';
               end
            else
               begin
               StSalDADS:= StSalDADS+')';
               StSalFI:= StSalFI+')';
               end;
            end;
      end;
//FIN PT7-1

   for i := 0 to TobEtabliss.Detail.Count - 1 do
       begin
       Etablissement:= TobEtabliss.Detail[i].GetValue('ET_ETABLISSEMENT');
       for j := 1 to 5 do
           begin
           StSQL:= 'SELECT D1.PDS_SALARIE,D1.PDS_DONNEEAFFICH'+
                   ' FROM DADSDETAIL D1'+
                   ' LEFT JOIN DADSDETAIL D2 ON'+
                   ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                   ' D1.PDS_TYPE=D2.PDS_TYPE AND' +
                   ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                   ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                   ' D1.PDS_DATEFIN=D2.PDS_DATEFIN AND'+
                   ' D1.PDS_EXERCICEDADS=D2.PDS_EXERCICEDADS WHERE'+
                   ' D1.PDS_SEGMENT="'+TabDads[j]+'" AND'+
                   ' D1.PDS_EXERCICEDADS="'+PGExercice+'" AND';
           if (GetControlText('CBTP')='X') then
              StSQL:= StSQL+' D1.PDS_TYPE<>"001" AND D1.PDS_TYPE<>"201" AND'
           else
              StSQL:= StSQL+' (D1.PDS_TYPE="001" OR D1.PDS_TYPE="201") AND';
           StSQL:= StSQL+' D2.PDS_SEGMENT="S41.G01.00.005" AND'+
                         ' D2.PDS_DONNEEAFFICH="'+Etablissement+'"'+StSalDADS;

          //DEB PT8
          if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
            StSQL := StSQL + ' AND D1.PDS_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
          //FIN PT8

           Q:= OpenSQL (StSQL, True);
           TobDADS:= Tob.Create('dads', nil, -1);
           TobDADS.LoadDetailDB('dads', '', '', Q, False);
           Ferme(Q);
           for x := 0 to TobDads.Detail.Count - 1 do
               begin
               Salarie:= TobDads.Detail[X].GetValue('PDS_SALARIE');
               T:= TobEtat.FindFirst(['PSA_SALARIE', 'PSA_ETABLISSEMENT'],
                                     [Salarie, Etablissement], False);
               if T <> nil then
                  begin
                  MtDads:= T.GetValue('DA' + IntToStr(j));
                  MtDads:= MtDads+StrToFloat(TobDads.Detail[x].GetValue('PDS_DONNEEAFFICH'));
                  T.PutValue('DA'+IntToStr(j), MtDads);
                  end
               else
                  begin
                  MtDads:= StrToFloat(TobDads.Detail[x].GetValue('PDS_DONNEEAFFICH'));
                  T:= Tob.Create('UneFille', TobEtat, -1);
                  T.AddchampSupValeur('PSA_SALARIE', Salarie, False);
                  T.AddchampSupValeur('PSA_ETABLISSEMENT', Etablissement, False);
                  T.AddchampSupValeur('PSA_LIBELLE', '', False);
                  T.AddchampSupValeur('PSA_PRENOM', '', False);
                  for a := 1 to 5 do
                      begin
                      if a = j then
                         T.AddchampSupValeur('DA' + IntToStr(a), MtDads, False)
                      else
                         T.AddchampSupValeur('DA' + IntToStr(a), 0, False);
                      T.AddchampSupValeur('FI' + IntToStr(a), 0, False);
                      T.AddchampSupValeur('EC' + IntToStr(a), 0, False);
                      end;
                  end;
               end;
           FreeAndNil (TobDads);
           end;
       for j:= 1 to 5 do
           begin
          StSQL := 'SELECT PHC_SALARIE,SUM(PHC_MONTANT) MTCUMUL' +
                       ' FROM HISTOCUMSAL WHERE'+
                       ' PHC_ETABLISSEMENT="'+Etablissement+'" AND'+
                       ' PHC_CUMULPAIE="'+TabCumul[j]+'" AND'+
                       ' PHC_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
                       ' PHC_DATEFIN<="'+UsDateTime(FinExer)+'"'+StSalFI;
          //DEB PT8
          if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
            StSQL := StSQL + ' AND PHC_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
          //FIN PT8

          StSQL := StSQL + ' GROUP BY PHC_SALARIE';
          Q:= OpenSQL(StSQL, True);
           TobCum:= Tob.Create('Lescumuls', nil, -1);
           TobCum.LoadDetailDB('Lescumuls', '', '', Q, False);
           Ferme(Q);
           for x := 0 to TobCum.Detail.Count - 1 do
               begin
               Salarie:= TobCum.Detail[x].GetValue('PHC_SALARIE');
               T:= TobEtat.FindFirst(['PSA_SALARIE', 'PSA_ETABLISSEMENT'],
                                     [Salarie, Etablissement], False);
               if T <> nil then
                  begin
                  MtCum:= Arrondi(TobCum.Detail[x].GetValue('MTCUMUL'), 0);
                  T.PutValue('FI' + IntToStr(j), MtCum);
                  end;
               end;
           FreeAndNil (TobCum);
           MoveCurProgressForm(TobEtabliss.Detail[i].GetValue('ET_LIBELLE'));
           end;
       end;
   FreeAndNil (TobEtabliss);
   for i:= 0 to TobEtat.Detail.Count - 1 do
       begin
       Q:= OpenSQL('SELECT PSA_LIBELLE,PSA_PRENOM'+
                   ' FROM SALARIES WHERE'+
                   ' PSA_SALARIE="'+TobEtat.Detail[i].GetValue('PSA_SALARIE')+'"', True);
       if not Q.Eof then
          begin
          TobEtat.Detail[i].PutValue('PSA_LIBELLE', Q.FindField('PSA_LIBELLE').AsString);
          TobEtat.Detail[i].PutValue('PSA_PRENOM', Q.FindField('PSA_PRENOM').AsString);
          end;
       Ferme(Q);
       for j := 1 to 5 do
           begin
           MtDads:= TobEtat.Detail[i].GetValue('DA' + IntToStr(j));
           MtCum:= TobEtat.Detail[i].GetValue('FI' + IntToStr(j));
           TobEtat.Detail[i].PutValue('EC' + IntToStr(j), MtDads - MtCum);
           end;
       end;
   end
else
   begin
{PT7-1
   if GetControlText('ETAB1') <> '' then
      StEtab:= ' AND PSA_ETABLISSEMENT>="'+GetControlText('ETAB1')+'" AND'+
               ' PSA_ETABLISSEMENT<="'+GetControlText('ETAB2')+'"'
   else
      StEtab := '';
   if GetControlText('SAL1') <> '' then
      StSal:= ' AND PSA_SALARIE>="'+GetControlText('SAL1')+'" AND'+
              ' PSA_SALARIE<="'+GetControlText('SAL2')+'"'
   else
      StSal := '';
}
   if (THMultiValCombobox(GetControl('ETAB')).Tous) then
      StEtab:= ' AND ET_FICTIF<>"X"'                //PT9
   else
      begin
      MCEtab:= GetControlText('ETAB');
      EtabSelect:= ReadTokenpipe(MCEtab, ';');
      StEtab:= ' AND (';
      While (EtabSelect <> '') do
            begin
            StEtab:= StEtab+' PSA_ETABLISSEMENT="'+EtabSelect+'"';
            EtabSelect := ReadTokenpipe(MCEtab,';');
            if (EtabSelect <> '') then
               StEtab:= StEtab+' OR'
            else
               StEtab:= StEtab+')';
            end;
      end;

   if (THMultiValCombobox(GetControl('SAL')).Tous) then
      StSal:= ''
   else
      begin
      MCSal:= GetControlText('SAL');
      Salarie:= ReadTokenpipe(MCSal, ';');
      StSal:= ' AND (';
      While (Salarie <> '') do
            begin
            StSal:= StSal+' PSA_SALARIE="'+Salarie+'"';
            Salarie := ReadTokenpipe(MCSal,';');
            if (Salarie <> '') then
               StSal:= StSal+' OR'
            else
               StSal:= StSal+')';
            end;
      end;
//FIN PT7-1

   StSQL:= 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT'+
           ' FROM SALARIES'+
           ' LEFT JOIN ETABLISS ON'+
           ' PSA_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
           ' (PSA_DATESORTIE>="'+UsdateTime(DebExer)+'" OR'+
           ' PSA_DATESORTIE="'+UsDateTime(IDate1900)+'") AND'+
           ' (PSA_SALARIE IN'+
           ' (SELECT DISTINCT (PDE_SALARIE) FROM DADSPERIODES WHERE'+
           ' PDE_EXERCICEDADS = "'+PGExercice+'" AND';
   if (GetControlText('CBTP')='X') then
      StSQL:= StSQL+' PDE_TYPE<>"001" AND PDE_TYPE<>"201"))'
   else
      StSQL:= StSQL+' (PDE_TYPE="001" OR PDE_TYPE="201")))';

  //DEB PT8
  Habilitation := '';
  if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
    Habilitation := ' AND ' + MonHabilitation.LeSQL;
  //FIN PT8

   StSQL:= StSQL+StEtab+Habilitation+StSal;
   Q:= OpenSQL(StSQL, True);
   TobEtat:= Tob.Create('Edition', nil, -1);
   TobEtat.LoadDetailDB('Edition', '', '', Q, False);
   Ferme(Q);
   InitMoveProgressForm (nil, 'Chargement des données',
                         'Veuillez patienter SVP ...', TobEtat.Detail.Count,
                         False, True);
   InitMove(TobEtat.Detail.Count, '');
   for j := 0 to TobEtat.Detail.Count - 1 do
       begin
       T:= Tobetat.Detail[j];
       Salarie:= TobEtat.Detail[j].GetValue('PSA_SALARIE');
       for i := 1 to 5 do
           begin
           StSQL:= 'SELECT PDS_SALARIE,PDS_DONNEEAFFICH'+
                   ' FROM DADSDETAIL WHERE'+
                   ' PDS_SEGMENT="'+TabDads[i]+'" AND'+
                   ' PDS_SALARIE="'+Salarie+'" AND'+
                   ' PDS_EXERCICEDADS = "'+PGExercice+'"';
           if (GetControlText('CBTP')='X') then
              StSQL:= StSQL+' AND PDS_TYPE<>"001" AND PDS_TYPE<>"201"'
           else
              StSQL:= StSQL+' AND (PDS_TYPE="001" OR PDS_TYPE="201")';
           Q:= OpenSQL(StSQL, True);
           TobDADS:= Tob.Create('dads', nil, -1);
           TobDADS.LoadDetailDB('dads', '', '', Q, False);
           Ferme(Q);
           MtDads:= 0;
           for x := 0 to TobDads.Detail.Count - 1 do
               MtDads:= MtDads+StrToFloat(TobDads.Detail[x].GetValue('PDS_DONNEEAFFICH'));
           FreeAndNil (TobDads);
           MtCum:= 0;
           Q:= OpenSQL('SELECT PHC_SALARIE,SUM(PHC_MONTANT) AS MTCUMUL'+
                       ' FROM HISTOCUMSAL WHERE'+
                       ' PHC_SALARIE="'+Salarie+'" AND'+
                       ' PHC_CUMULPAIE="'+TabCumul[i]+'" AND'+
                       ' PHC_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
                       ' PHC_DATEFIN<="'+UsDateTime(FinExer)+'"'+
                       ' GROUP BY PHC_SALARIE', True);
           if not Q.Eof then
              MtCum:= Arrondi(Q.FindField('MTCUMUL').AsFloat, 0);
           Ferme(Q);
           T.AddchampSupValeur('DA'+IntToStr(i), MtDads, False);
           T.AddchampSupValeur('FI'+IntToStr(i), MtCum, False);
           T.AddchampSupValeur('EC'+IntToStr(i), MtDads-MtCum, False);
           end;
       MoveCurProgressForm('Salarie : ' + TobEtat.Detail[j].GetValue('PSA_LIBELLE'));
       end;
   end;
FiniMoveProgressForm;
if RuptEtablissement = True then
   begin
   if GetCheckBoxState('TRISAL') = CbChecked then
      TobEtat.Detail.Sort('PSA_ETABLISSEMENT;PSA_LIBELLE;PSA_SALARIE')
   else
      TobEtat.Detail.Sort('PSA_ETABLISSEMENT;PSA_SALARIE');
   end
else
   begin
   if GetCheckBoxState('TRISAL') = CbChecked then
      TobEtat.Detail.Sort('PSA_LIBELLE;PSA_SALARIE')
   else
      TobEtat.Detail.Sort('PSA_SALARIE');
   end;

TED:= TobEtat.FindFirst([''], [''], TRUE);
Count:= TobEtat.Detail.Count;
for i := 0 to Count - 1 do
    begin
    if (TED <> nil) then
       begin
       Difference:= False;
       for j := 1 to 5 do
           begin
           if ((Abs(TED.GetValue('EC'+IntToStr(j)))>1) or
              (Abs(TED.GetValue('EC'+IntToStr(j)))<-1)) then
              Difference := True;
           end;
       if Difference = False then
          FreeAndNil(TED);
       TED:= TobEtat.FindNext([''], [''], TRUE);
       end;
    end;
{if TobEtat.Detail.Count <= 0 then   //PT5 Mis en commentaire
   begin
   PGIBox('Aucune différence à éditer', Ecran.Caption);
   FreeAndNil(TobEtat);
   Exit;
   end;}
StPages := AglGetCriteres (Pages, FALSE);
TFQRS1(Ecran).TypeEtat:= 'E';
TFQRS1(Ecran).NatureEtat:= 'PDA';
TFQRS1(Ecran).CodeEtat:= 'PCF';
TFQRS1(Ecran).LaTob:= TobEtat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/03/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITCOMPDADS.OnClose ;
begin
Inherited ;
FreeAndNil (TobEtat);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITCOMPDADS.CocheSansDetail(Sender : TObject);
begin
If GetCheckBoxState('CDETAILSAL') = CbChecked then
   begin
   SetControlEnabled('TRISAL',False);
{PT7-1
   SetControlEnabled('SAL1',False);
   SetControlEnabled('SAL2',False);
}
   SetControlEnabled('SAL',False);
   end
else
   begin
   SetControlEnabled('TRISAL',True);
{PT7-1
   SetControlEnabled('SAL1',True);
   SetControlEnabled('SAL2',True);
}
   SetControlEnabled('SAL',True);
   end;
end;


//PT7-2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITCOMPDADS.DateDecaleChange(Sender: TObject);
begin
AglLanceFiche ('PAY', 'DADS_DATE', '', '', '');
SetControlText ('DATEDEB', DateToStr(DebExer));
SetControlText ('DATEFIN', DateToStr(FinExer));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITCOMPDADS.Parametrage(Sender: TObject);
var
StExer : string;
QExer : TQuery;
begin
if (GetCheckBoxState('CBTP')=CbChecked) then
   begin
   Car.Enabled:= True;
   Car.Visible:= True;
   SetControlVisible ('LCAR', True);
   end
else
   begin
   Car.Enabled:= False;
   Car.Visible:= False;
   SetControlVisible ('LCAR', False);
   end;
   
PGAnnee:= Annee.value;
StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';
QExer:= OpenSQL(StExer,TRUE) ;
if (NOT QExer.EOF) then
   PGExercice:= QExer.FindField ('PEX_ANNEEREFER').AsString
else
   begin
   PGExercice:= '';
   Ferme (QExer);
   exit;
   end;

if (GetCheckBoxState('CBTP')=CbChecked) then
   begin
   if ((Car.Value>='M00') and (Car.Value <='M99')) then
      begin
      TypeD:= '005';
      if (Car.Value = 'M01') then
         begin
         DebExer:= StrToDate('01/01/'+PGExercice);
         FinExer:= StrToDate('31/01/'+PGExercice);
         end
      else
      if (Car.Value = 'M02') then
         begin
         DebExer:= StrToDate('01/02/'+PGExercice);
         FinExer:= FinDeMois(StrToDate('28/02/'+PGExercice));
         end
      else
      if (Car.Value = 'M03') then
         begin
         DebExer:= StrToDate('01/03/'+PGExercice);
         FinExer:= StrToDate('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'M04') then
         begin
         DebExer:= StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/04/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M05') then
         begin
         DebExer:= StrToDate('01/05/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/05/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M06') then
         begin
         DebExer:= StrToDate('01/06/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M07') then
         begin
         DebExer:= StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/07/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M08') then
         begin
         DebExer:= StrToDate('01/08/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/08/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M09') then
         begin
         DebExer:= StrToDate('01/09/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M10') then
         begin
         DebExer:= StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/10/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M11') then
         begin
         DebExer:= StrToDate('01/11/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/11/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M12') then
         begin
         DebExer:= StrToDate('01/12/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
         end;
      end;
   if ((Car.Value>='T00') and (Car.Value <='T99')) then
      begin
      TypeD:= '004';
      if (Car.Value = 'T01') then
         begin
{PT10
         DebExer:= StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer:= StrToDate('01/01/'+PGExercice);
         FinExer:= StrToDate('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'T02') then
         begin
{
         DebExer:= StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer:= StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'T03') then
         begin
{
         DebExer:= StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer:= StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'T04') then
         begin
{
         DebExer:= StrToDate('01/01/'+PGExercice);
         FinExer:= StrToDate('31/03/'+PGExercice);
}
         DebExer:= StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer:= StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
//FIN PT10         
         end;
      end;
   if ((Car.Value>='S00') and (Car.Value <='S99')) then
      begin
      PGIBox ('Périodicité interdite', Ecran.Caption);
      Car.Value:= 'A00';
      end;
   if Car.Value='A00' then
      begin
      TypeD:= '002';
      DebExer:= StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
      FinExer:= StrToDate('31/03/'+PGExercice);
      end;
   end
else
   begin
   TypeD:= '001';
   if NOT QExer.eof then
      begin
      DebExer:= QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
      FinExer:= QExer.FindField ('PEX_DATEFIN').AsDateTime;
      end;
   end;
Ferme(QExer);

SetControlText ('DATEDEB', DateToStr(DebExer));
SetControlText ('DATEFIN', DateToStr(FinExer));
end;
//FIN PT7-2

initialization
  registerclasses([TOF_PGEDITCOMPDADS]);
end.

