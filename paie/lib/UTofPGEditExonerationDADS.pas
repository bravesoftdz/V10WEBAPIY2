{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 19/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EDIT_EXODADS ()
Mots clefs ... : TOF;EXONERATION DADS-U
*****************************************************************}
{
PT1   : 01/02/2005 VG V_60 Edition en CWAS
PT2   : 02/03/2005 VG V_60 LanceEtatTob propriété de QRS1 - FQ N°11955
PT3   : 09/01/2006 PH V_65 FQ 12281 Multi-période DADS_U
PT4-1 : 17/10/2006 VG V_70 Remplacement des ellipsis et combo par des multi-val
                           combo
PT4-2 : 17/10/2006 VG V_70 Gestion du décalage de paie - FQ N°12860
PT5   : 30/01/2007 FC V_80 Mise en place filtrage des habilitations/poupulations
PT6   : 26/09/2007 VG V_80 Adaptation cahier des charges V8R05
PT7   : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06
PT8   : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT9   : 30/11/2007 VG V_80 Chargement de PGExercice - FQ N°14888
}
unit UTofPGEditExonerationDADS;

interface
uses StdCtrls,
     Controls,
     Classes,
     sysutils,
     ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,
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
     HQRY,
     HTB97,
     PgDADSCommun,
     P5Def
     ;

type
  TOF_PGEDITEXONERATIONDADS = class(TOF)
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override ;
  private
    TobEtat : TOB;
    procedure CocheSansDetail(Sender : TObject);
    procedure DateDecaleChange(Sender: TObject);
    procedure Parametrage(Sender: TObject);
  end;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 19/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITEXONERATIONDADS.OnArgument;
var
AnneeE, AnneePrec, MoisE, ComboExer, StPlus: string;
JourJ: TDateTime;
AnneeA, Jour, MoisM: Word;
Check : TCheckBox;
ModifDateDecale : TToolbarButton97;
Annee : THValComboBox;
begin
inherited;
{PT4-1
DateDeb:= IDate1900;
DateFin:= IDate1900;
RecupMinMaxTablette ('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
SetControlText ('ETAB1', Min);
SetControlText ('ETAB2', Max);
RecupMinMaxTablette ('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
SetControlText ('Sal1', Min);
SetControlText ('Sal2', Max);
}
JourJ:= Date;
DecodeDate (JourJ, AnneeA, MoisM, Jour);
if MoisM>9 then
   AnneePrec:= IntToStr(AnneeA)
else
   AnneePrec:= IntToStr(AnneeA-1);

{PT4-1
if RendExerSocialPrec(MoisE, AnneeE, ComboExer, Deb, Fin, AnneePrec)=TRUE then
}
if RendExerSocialPrec(MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec)=TRUE then
   begin
   SetControlText('ANNEE', ComboExer);
   PGExercice:= RechDom ('PGANNEESOCIALE', GetControlText('ANNEE'), False);       //PT9
   end
else
   PGIBox('L''exercice ' + AnneePrec + ' n''existe pas', Ecran.Caption);
{PT4-1
Edit:= THEdit(GetControl('SAL1'));
if Edit<>nil then
   Edit.OnExit:= ExitEdit;
Edit:= THEdit(GetControl('SAL2'));
if Edit<>nil then
   Edit.OnExit:= ExitEdit;
}   
SetControlText ('DOSSIER', GetParamSoc('SO_LIBELLE'));
Check:= TCheckBox(GetControl('CDETAILSAL'));
If Check<>Nil then
   Check.OnClick:= CocheSansDetail;

SetControlText('DATEDEB', DateToStr(DebExer));
SetControlText('DATEFIN', DateToStr(FinExer));

//PT4-2
// Gestion du bouton de modification des dates
ModifDateDecale:= TToolbarButton97(GetControl('BMODIFDATE'));
if ModifDateDecale<>nil then
   ModifDateDecale.OnClick:= DateDecaleChange;

Annee:= THValComboBox (GetControl ('ANNEE'));
If Annee <> NIL then
   Annee.OnChange := Parametrage;
//FIN PT4-2

//PT8
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('ETAB', 'Plus', StPlus);
//FIN PT8
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 19/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITEXONERATIONDADS.OnUpdate;
var
Q : TQuery;
TobExo, TobSalarie, TE, TS : Tob;
TabMotifExo : array[1..7] of string;
i, x : Integer;
Salarie : string;
Pages : TPageControl;
EtabSelect, MCEtab, MCSal, StCriteres, StEtab, StPages, StSal, StWhere : string;
Montant : Integer;
Habilitation:String;
begin
inherited;
{PT4-1
DateDeb:= IDate1900;
DateFin:= IDate1900;
}   
Pages:= TPageControl(GetControl('PAGES'));
StCriteres:= RecupWhereCritere(Pages);

{PT4-1
Q:= OpenSQL ('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL WHERE'+
             ' PEX_EXERCICE="' + GetControlText('ANNEE') + '"', True);
if not Q.eof then
   begin
   DateDeb:= Q.FindField('PEX_DATEDEBUT').AsDateTime;
   DateFin:= Q.FindField('PEX_DATEFIN').AsDateTime;
   end;
Ferme(Q);
SetControlText('DATEDEB', DateToStr(DateDeb));
SetControlText('DATEFIN', DateToStr(DateFin));
}

//Récupération des motifs d'éxonération sélectionnés pour l'édition
//DEB PT5
Habilitation := '';
if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
begin
  Habilitation := ' AND PDS_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
end;
//FIN PT5

if StCriteres<>'' then
   StWhere:= StCriteres+'AND PDS_SEGMENT="S41.G01.06.001" '
else
   StWhere:= 'WHERE PDS_SEGMENT="S41.G01.06.001"';
StWhere := StWhere + Habilitation; //PT5
for i:=1 to 7 do
    begin
    TabMotifExo[i]:= '';
    SetControlText ('MOTIXEXO'+IntToStr(i), '');
    end;
Q:= OpenSQL ('SELECT PDS_DONNEEAFFICH,COUNT (PDS_SALARIE) NBTYPEEXO FROM'+
             ' DADSDETAIL '+StWhere+'GROUP BY PDS_DONNEEAFFICH '+
             'ORDER BY NBTYPEEXO DESC', True);
TobExo:= Tob.Create('Lesexoneration', nil, -1);
TobExo.LoadDetailDB('Lesexoneration', '', '', Q, False);
Ferme(Q);
{Récupération des 7 motifs les plus fréquents (si plus de 7 sélectionnés) car
affichage en colonne}
for i:=0 to TobExo.Detail.Count-1 do
    begin
    if i>6 then
       break;
    TabMotifExo[i + 1]:= TobExo.Detail[i].GetValue('PDS_DONNEEAFFICH');
    if TobExo.Detail[i].GetValue('PDS_DONNEEAFFICH')<>'' then
       begin
       SetControlText ('MOTIXEXO'+IntToStr(i+1),
                       RechDom('PGEXONERATION',
                       TobExo.Detail[i].GetValue('PDS_DONNEEAFFICH'), False));
       SetControlText ('CODEEXO'+IntToStr(i+1),
                       TobExo.Detail[i].GetValue('PDS_DONNEEAFFICH'));
       end
    else
       begin
       SetControlText('MOTIXEXO'+IntToStr(i+1), '');
       SetControlText('CODEEXO'+IntToStr(i+1), '');
       end;
    end;
FreeAndNil (TobExo);
{PT4-1
if GetControlText('ETAB1')<>'' then
   StEtab:= ' AND PSA_ETABLISSEMENT>="'+GetControlText('ETAB1')+'" AND'+
            ' PSA_ETABLISSEMENT<="'+GetControlText('ETAB2')+'"'
else
   StEtab:= '';

if GetControlText('SAL1')<>'' then
   StSal:= ' AND PSA_SALARIE>="'+GetControlText('SAL1')+'" AND PSA_SALARIE<="'+
           GetControlText('SAL2')+'"'
else
   StSal:= '';
}
if (THMultiValCombobox(GetControl('ETAB')).Tous) then
   StEtab:= ' AND ET_FICTIF<>"X"'                //PT8
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
//FIN PT4-1

InitMoveProgressForm (nil, 'Chargement des données',
                      'Veuillez patienter SVP ...', 7, False, True);
InitMove(7, '');
FreeAndNil (TobEtat);
TobEtat := Tob.Create('Edition', nil, -1);
// Calcul des montants par motifs
for i:=1 to 7 do
    begin
    if TabMotifExo[i]='' then
       break;
    //DEB PT5
    Habilitation := '';
    if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
    begin
      Habilitation := ' AND ' + MonHabilitation.LeSQL;
    end;
    //FIN PT5
// Calcul des bases bruts
    Q:= OpenSQL ('SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM,'+
                 ' PSA_ETABLISSEMENT, D2.PDS_SALARIE,'+
                 ' (D2.PDS_DONNEEAFFICH) MONTANT'+
                 ' FROM SALARIES'+
                 ' LEFT JOIN DADSDETAIL D1 ON'+
                 ' PSA_SALARIE=D1.PDS_SALARIE'+
                 ' LEFT JOIN DADSDETAIL D2 ON'+
                 ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                 ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                 ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                 ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                 ' D1.PDS_DATEFIN=D2.PDS_DATEFIN AND'+
                 ' D1.PDS_EXERCICEDADS=D2.PDS_EXERCICEDADS'+
                 ' LEFT JOIN ETABLISS ON'+
                 ' PSA_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
                 ' D1.PDS_SEGMENT="S41.G01.06.001" AND'+
                 ' D2.PDS_SEGMENT="S41.G01.06.002.001" AND'+
                 ' D1.PDS_EXERCICEDADS="'+PGExercice+'" AND'+
                 ' D1.PDS_DONNEEAFFICH="'+TabMotifExo[i]+'" AND'+
                 ' ((D1.PDS_ORDRESEG="601" AND D2.PDS_ORDRESEG="602") OR'+
                 ' (D1.PDS_ORDRESEG="606" AND D2.PDS_ORDRESEG="607") OR'+
                 ' (D1.PDS_ORDRESEG="611" AND D2.PDS_ORDRESEG="612") OR'+
                 ' (D1.PDS_ORDRESEG="616" AND D2.PDS_ORDRESEG="617") OR'+
                 ' (D1.PDS_ORDRESEG="621" AND D2.PDS_ORDRESEG="622") OR'+
                 ' (D1.PDS_ORDRESEG="626" AND D2.PDS_ORDRESEG="627") OR'+
                 ' (D1.PDS_ORDRESEG="631" AND D2.PDS_ORDRESEG="632") OR'+
                 ' (D1.PDS_ORDRESEG="636" AND D2.PDS_ORDRESEG="637") OR'+
                 ' (D1.PDS_ORDRESEG="641" AND D2.PDS_ORDRESEG="642") OR'+
                 ' (D1.PDS_ORDRESEG="646" AND D2.PDS_ORDRESEG="647")) '+StSal+
                 ' '+StEtab + Habilitation, True);    //PT5
    TobSalarie:= Tob.Create('LesExosalarie', nil, -1);
    TobSalarie.LoadDetailDB('LesExosalarie', '', '', Q, False);
    Ferme(Q);
    for x:=0 to TobSalarie.Detail.Count-1 do
        begin
        Salarie:= TobSalarie.Detail[x].GetValue('PSA_SALARIE');
        if IsNumeric(TobSalarie.Detail[x].GetValue('MONTANT')) then
           Montant:= StrToInt(TobSalarie.Detail[x].GetValue('MONTANT'))
        else
           Montant:= 0;
        TE:= TobEtat.FindFirst(['PSA_SALARIE'], [Salarie], False);
        if TE<>nil then
           TE.PutValue ('BRUT'+IntToStr(i),
                        TE.GetValue('BRUT'+IntToStr(i))+Montant)
        else
           begin
           TE:= Tob.Create('FilleEtat', TobEtat, -1);
           TE.AddchampSupValeur ('PSA_SALARIE', Salarie, False);
           TE.AddchampSupValeur ('PSA_ETABLISSEMENT',
                                 TobSalarie.Detail[x].GetValue('PSA_ETABLISSEMENT'),
                                 False);
           TE.AddchampSupValeur ('PSA_LIBELLE',
                                 TobSalarie.Detail[x].GetValue('PSA_LIBELLE'),
                                 False);
           TE.AddchampSupValeur ('PSA_PRENOM',
                                 TobSalarie.Detail[x].GetValue('PSA_PRENOM'),
                                 False);
           TE.AddchampSupValeur ('BRUT1', 0, False);
           TE.AddchampSupValeur ('BRUT2', 0, False);
           TE.AddchampSupValeur ('BRUT3', 0, False);
           TE.AddchampSupValeur ('BRUT4', 0, False);
           TE.AddchampSupValeur ('BRUT5', 0, False);
           TE.AddchampSupValeur ('BRUT6', 0, False);
           TE.AddchampSupValeur ('BRUT7', 0, False);
           TE.AddchampSupValeur ('PLAF1', 0, False);
           TE.AddchampSupValeur ('PLAF2', 0, False);
           TE.AddchampSupValeur ('PLAF3', 0, False);
           TE.AddchampSupValeur ('PLAF4', 0, False);
           TE.AddchampSupValeur ('PLAF5', 0, False);
           TE.AddchampSupValeur ('PLAF6', 0, False);
           TE.AddchampSupValeur ('PLAF7', 0, False);
           TE.AddchampSupValeur ('NBPERIODES', 0, False);
           TE.PutValue('BRUT'+IntToStr(i), Montant);
           end;
        end;
    FreeAndNil (TobSalarie);

    //DEB PT5
    Habilitation := '';
    if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
    begin
      Habilitation := ' AND ' + MonHabilitation.LeSQL;
    end;
    //FIN PT5
//Calcul des bases plafonnés
    Q:= OpenSQL ('SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM,'+
                 ' PSA_ETABLISSEMENT,D2.PDS_SALARIE,'+
                 '(D2.PDS_DONNEEAFFICH) MONTANT'+
                 ' FROM SALARIES'+
                 ' LEFT JOIN DADSDETAIL D1 ON'+
                 ' PSA_SALARIE=D1.PDS_SALARIE'+
                 ' LEFT JOIN DADSDETAIL D2 ON'+
                 ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                 ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                 ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                 ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                 ' D1.PDS_DATEFIN=D2.PDS_DATEFIN AND'+
                 ' D1.PDS_EXERCICEDADS=D2.PDS_EXERCICEDADS'+
                 ' LEFT JOIN ETABLISS ON'+
                 ' PSA_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
                 ' D1.PDS_SEGMENT="S41.G01.06.001" AND'+
                 ' D2.PDS_SEGMENT="S41.G01.06.003.001" AND'+
                 ' D1.PDS_EXERCICEDADS="'+PGExercice+'" AND'+
                 ' D1.PDS_DONNEEAFFICH="'+TabMotifExo[i]+'" AND'+
                 ' ((D1.PDS_ORDRESEG="601" AND D2.PDS_ORDRESEG="604") OR'+
                 ' (D1.PDS_ORDRESEG="606" AND D2.PDS_ORDRESEG="609") OR'+
                 ' (D1.PDS_ORDRESEG="611" AND D2.PDS_ORDRESEG="614") OR'+
                 ' (D1.PDS_ORDRESEG="616" AND D2.PDS_ORDRESEG="619") OR'+
                 ' (D1.PDS_ORDRESEG="621" AND D2.PDS_ORDRESEG="624") OR'+
                 ' (D1.PDS_ORDRESEG="626" AND D2.PDS_ORDRESEG="629") OR'+
                 ' (D1.PDS_ORDRESEG="631" AND D2.PDS_ORDRESEG="634") OR'+
                 ' (D1.PDS_ORDRESEG="636" AND D2.PDS_ORDRESEG="639") OR'+
                 ' (D1.PDS_ORDRESEG="641" AND D2.PDS_ORDRESEG="644") OR'+
                 ' (D1.PDS_ORDRESEG="646" AND D2.PDS_ORDRESEG="649"))'+StSal+
                 ' '+StEtab + Habilitation, True); //PT5
    TobSalarie:= Tob.Create('LesExosalarie', nil, -1);
    TobSalarie.LoadDetailDB('LesExosalarie', '', '', Q, False);
    Ferme(Q);
    for x:=0 to TobSalarie.Detail.Count-1 do
        begin
        Salarie:= TobSalarie.Detail[x].GetValue('PSA_SALARIE');
        if IsNumeric(TobSalarie.Detail[x].GetValue('MONTANT')) then
           Montant:= StrToInt(TobSalarie.Detail[x].GetValue('MONTANT'))
        else
           Montant := 0;
        TE:= TobEtat.FindFirst(['PSA_SALARIE'], [Salarie], False);
        if TE<>nil then
           TE.PutValue ('PLAF'+IntToStr(i),
                        TE.GetValue('PLAF'+IntToStr(i))+Montant)
        else
           begin
           TE:= Tob.Create('FilleEtat', TobEtat, -1);
           TE.AddchampSupValeur ('PSA_SALARIE', Salarie, False);
           TE.AddchampSupValeur ('PSA_ETABLISSEMENT',
                                 TobSalarie.Detail[x].GetValue('PSA_ETABLISSEMENT'),
                                 False);
           TE.AddchampSupValeur ('PSA_LIBELLE',
                                 TobSalarie.Detail[x].GetValue('PSA_LIBELLE'),
                                 False);
           TE.AddchampSupValeur ('PSA_PRENOM',
                                 TobSalarie.Detail[x].GetValue('PSA_PRENOM'),
                                 False);
           TE.AddchampSupValeur ('BRUT1', 0, False);
           TE.AddchampSupValeur ('BRUT2', 0, False);
           TE.AddchampSupValeur ('BRUT3', 0, False);
           TE.AddchampSupValeur ('BRUT4', 0, False);
           TE.AddchampSupValeur ('BRUT5', 0, False);
           TE.AddchampSupValeur ('BRUT6', 0, False);
           TE.AddchampSupValeur ('BRUT7', 0, False);
           TE.AddchampSupValeur ('PLAF1', 0, False);
           TE.AddchampSupValeur ('PLAF2', 0, False);
           TE.AddchampSupValeur ('PLAF3', 0, False);
           TE.AddchampSupValeur ('PLAF4', 0, False);
           TE.AddchampSupValeur ('PLAF5', 0, False);
           TE.AddchampSupValeur ('PLAF6', 0, False);
           TE.AddchampSupValeur ('PLAF7', 0, False);
           TE.AddchampSupValeur ('NBPERIODES', 0, False);
           TE.PutValue('PLAF'+IntToStr(i), Montant);
           end;
        end;
    FreeAndNil (TobSalarie);
    MoveCurProgressForm ('Type exonération : '+
                         RechDom('PGEXONERATION', TabMotifExo[i], False));
    end;

FiniMoveProgressForm;
//DEB PT5
Habilitation := '';
if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
begin
  Habilitation := ' AND PDE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
end;
//FIN PT5
Q:= OpenSQL ('SELECT PDE_SALARIE, COUNT(PDE_ORDRE) NBPERIODES'+
             ' FROM DADSPERIODES WHERE'+
             ' PDE_ORDRE>0 AND'+
             ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
             Habilitation + {PT5}
             ' GROUP BY PDE_SALARIE', True);
TobSalarie:= Tob.Create('Compteperiodes', nil, -1);
TobSalarie.LoadDetailDB('Compteperiodes', '', '', Q, False);
Ferme(Q);
for i:=0 to TobEtat.Detail.Count-1 do
    begin
    TS:= TobSalarie.FindFirst (['PDE_SALARIE'],
                               [TobEtat.Detail[i].GetValue('PSA_SALARIE')],
                               False);
    if TS<>nil then
       begin
       TobEtat.Detail[i].PutValue('NBPERIODES', TS.GetValue('NBPERIODES'));
       FreeAndNil (TS);
       end;
    end;
FreeAndNil (TobSalarie);

if GetCheckBoxState('CRUPTETAB')=CbChecked then
   begin
   if GetCheckBoxState('TRISAL')=CbChecked then
      TobEtat.Detail.Sort('PSA_ETABLISSEMENT;PSA_LIBELLE;PSA_SALARIE')
   else
      TobEtat.Detail.Sort('PSA_ETABLISSEMENT;PSA_SALARIE');
   end
else
   begin
   if GetCheckBoxState('TRISAL')=CbChecked then
      TobEtat.Detail.Sort('PSA_LIBELLE;PSA_SALARIE')
   else
      TobEtat.Detail.Sort('PSA_SALARIE');
   end;

StPages:= AglGetCriteres (Pages, FALSE);
TFQRS1(Ecran).TypeEtat:= 'E';
TFQRS1(Ecran).NatureEtat:= 'PDA';
TFQRS1(Ecran).CodeEtat:= 'PEX';
TFQRS1(Ecran).LaTob:= TobEtat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/03/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITEXONERATIONDADS.OnClose ;
begin
Inherited ;
FreeAndNil (TobEtat);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 19/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGEDITEXONERATIONDADS.CocheSansDetail(Sender : TObject);
begin
If GetCheckBoxState('CDETAILSAL') = CbChecked then
   begin
   SetControlEnabled('TRISAL',False);
{PT4-1
   SetControlEnabled('SAL1',False);
   SetControlEnabled('SAL2',False);
}
   SetControlEnabled('SAL',False);
   end
else
   begin
   SetControlEnabled('TRISAL',True);
{
   SetControlEnabled('SAL1',True);
   SetControlEnabled('SAL2',True);
}
   SetControlEnabled('SAL',True);
   end;
//FIN PT4-1
end;

//PT4-2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITEXONERATIONDADS.DateDecaleChange(Sender: TObject);
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
procedure TOF_PGEDITEXONERATIONDADS.Parametrage(Sender: TObject);
var
Q: TQuery;
begin
Q := OpenSQL('SELECT PEX_DATEDEBUT,PEX_DATEFIN, PEX_ANNEEREFER'+
             ' FROM EXERSOCIAL WHERE'+
             ' PEX_EXERCICE="' + GetControlText('ANNEE') + '"', True);
if not Q.eof then
   begin
   DebExer:= Q.FindField ('PEX_DATEDEBUT').AsDateTime;
   FinExer:= Q.FindField ('PEX_DATEFIN').AsDateTime;
   PGExercice:= Q.FindField ('PEX_ANNEEREFER').AsString; //PT7
   end;
Ferme (Q);

SetControlText ('DATEDEB', DateToStr(DebExer));
SetControlText ('DATEFIN', DateToStr(FinExer));
end;
//FIN PT4-2

initialization
  registerclasses([TOF_PGEDITEXONERATIONDADS]);
end.

