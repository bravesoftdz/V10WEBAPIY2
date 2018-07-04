{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PGEDITDADSHON ()
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************
PT1   : 19/02/2002 JL V571 Correction requête pour Oracle, rajout de la
                           condition CEtab2 dans CEtab1
PT2   : 27/02/2002 JL V571 Correction requête pour Oracle, suppression de CEtab1
                           et CEtab2
PT3   : 26/04/2002 VG V571 Par défaut, année précédente jusqu'au mois de
                           septembre
                           A partir du mois d'octobre, année en cours (pour
                           tests)
PT4   : 29/01/2003 VG V591 Affichage d'un message si l'exercice par défaut
                           n'existe pas - FQ N°10469
PT5   : 14/04/2004 VG V_50 Ajout de l'année de référence au niveau de l'état
PT6   : 01/03/2005 VG V_60 On passe désormais par une tob pour éditer l'état
PT7   : 22/11/2005 JL V_65 12718 Affichage donnée fichier pour code modal prise en charge
                                 + affichage date fin exer compta si 31/12
PT8-1 : 16/10/2006 VG V_70 Remplacement des ellipsis et combo par des multi-val
                           combo
PT8-2 : 16/10/2006 VG V_70 Lancement de l'état du dernier exercice connu même si
                           on se positionne sur un exercice posterieur
PT9   : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
}
Unit PGEDITDADSHON_TOF ;

Interface
uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     Controls,Classes,sysutils,UTOB,paramsoc,ed_tools,
{$IFDEF EAGLCLIENT}
     eQRS1,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,HStatus,UTOF;

Type
  TOF_PGEDITDADSHON = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    PRIVATE
    THAnneeRef:THValCombobox;
    TobPeriodes : Tob;
    procedure ChargerLexiqueDads (var TobLexique:tob; Annee:String);
  end ;

Implementation

uses PgOutils2,
     PgDADSCommun,
     PgDADSOutils;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.OnNew ;
begin
Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.OnDelete ;
begin
Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.OnUpdate ;
var
Annee, AnneeCompta, AnneeF, CodeEtat, EtabSelect, MaxModele, MCEtab, Nature, OrderBy : String;
OrdreSeg, OrdreSegS, Req, requete, Segment, StEtab, StModele, Tablette : String;
ValeurSeg : String;
Q, QExer, QModele : TQuery;
i, NumOrdre, OrdreSegI, x : integer;
T, TD, TEtabDADS, TEtabDADSD, TobDetail, TobLexique : TOB;
begin
Inherited ;
FreeAndNil (TobPeriodes);
Annee:= THAnneeRef.Value;
AnneeCompta := RechDom('PGANNEESOCIALE', GetControlText('ANNEE'), False);
AnneeF:= Copy (RechDom('PGANNEESOCIALE', GetControlText('ANNEE'), False),4,1);
DebExer:= IDate1900;
FinExer:= IDate1900;
QExer:= OpenSQL ('SELECT PEX_DATEDEBUT,PEX_DATEFIN'+
                 ' FROM EXERSOCIAL WHERE'+
                 ' PEX_EXERCICE="'+Annee+'"',True);
If not QExer.eof then
   begin
   DebExer:= QExer.FindField('PEX_DATEDEBUT').AsDateTime;
   FinExer:= QExer.FindField('PEX_DATEFIN').AsDateTime;
   end
Else
   PGIBox ('Il n''existe pas d''exercice pour l''année '+Annee,Ecran.Caption);
Ferme(QExer);
{PT8-1
Etab1:= GetControlText('ETAB1');
Etab2:= GetControlText('ETAB2');
}
SetControlText ('DATEDEB',DateTimeToStr(DebExer));
SetControlText ('DATEFIN',DateTimeToStr(FinExer));

if (THMultiValCombobox(GetControl('ETAB')).Tous) then
{PT9
   StEtab:= ''
}
   begin
   StEtab:= ' AND (';
   TEtabDADS:= TOB.Create ('Les etablissements', nil, -1);
   ChargeEtabNonFictif (TEtabDADS);
   if (TEtabDADS<>nil) then
      begin
      TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
      if (TEtabDADSD<>nil) then
         EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
      else
         StEtab:= '';
      while (EtabSelect<>'') do
            begin
            StEtab:= StEtab+' PDS_DONNEEAFFICH="'+EtabSelect+'"';
            TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
            if (TEtabDADSD<>nil) then
               EtabSelect:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
            else
               EtabSelect:='';
            if (EtabSelect<>'') then
               StEtab:= StEtab+' OR'
            else
               StEtab:= StEtab+')';
            end;
      end
   else
      StEtab:= '';
   FreeAndNil (TEtabDADS);
   end
//FIN PT9
else
   begin
   MCEtab:= GetControlText('ETAB');
   EtabSelect:= ReadTokenpipe(MCEtab, ';');
   StEtab:= ' AND (';
   While (EtabSelect <> '') do
         begin
         StEtab:= StEtab+' PDS_DONNEEAFFICH="'+EtabSelect+'"';
         EtabSelect := ReadTokenpipe(MCEtab,';');
         if (EtabSelect <> '') then
            StEtab:= StEtab+' OR'
         else
            StEtab:= StEtab+')';
         end;
   end;
//FIN PT8-1

requete:= 'SELECT PDE_ORDRE,PDS_ORDRESEG,PDS_DONNEEAFFICH'+
          ' FROM DADSPERIODES'+
          ' LEFT JOIN DADSDETAIL ON'+
          ' PDS_ORDRE=PDE_ORDRE AND'+
          ' PDE_SALARIE=PDS_SALARIE WHERE'+
          ' PDE_SALARIE LIKE "--H%"'+StEtab+' AND'+
          ' PDS_SEGMENT="S70.G01.00.014" AND'+
          ' PDS_ORDRE IN (SELECT PDS_ORDRE FROM DADSDETAIL WHERE'+
          ' PDS_SEGMENT="S70.G01.01.002" AND'+
          ' PDS_SALARIE LIKE "--H%" AND'+
          ' (PDS_TYPE="001" OR PDS_TYPE="201") AND'+
          ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
          ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'")';
OrderBy:= ' ORDER BY PDE_ORDRE';
Q:= OpenSQL(requete,True);
TobPeriodes:= Tob.Create('LesPeriodes',Nil,-1);
TobPeriodes.LoadDetailDB('LesPeriodes','','',Q,False);
Ferme(Q);

TobLexique:= Tob.Create('LesSegments',Nil,-1);
ChargerLexiqueDads(TobLexique, AnneeF);

InitMoveProgressForm (NIL,'Chargement des données',
                     'Veuillez patienter SVP ...', TobPeriodes.Detail.Count,
                     False,True);
InitMove(TobPeriodes.Detail.Count,'');

For i:=0 to TobPeriodes.Detail.Count-1 do
    begin
    T:= TobPeriodes.Detail[i];
    T.AddchampSupValeur('DOSSIER', GetControlText('DOSSIER'), False);
    T.AddchampSupValeur('DATEDEB', GetControlText('DATEDEB'), False);
    T.AddchampSupValeur('DATEFIN', GetControlText('DATEFIN'), False);
    NumOrdre:= T.GetValue('PDE_ORDRE');
    Req:= 'SELECT *'+
          ' FROM DADSDETAIL WHERE'+
          ' PDS_SALARIE LIKE "--H%" AND'+
          ' PDS_ORDRE='+IntToStr(NumOrdre)+' AND'+
          ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
          ' PDS_DATEDEBUT<="'+UsDateTime(FinExer)+'" AND'+
          ' (PDS_TYPE="001" OR PDS_TYPE="201")'+
          ' ORDER BY PDS_ORDRE';
    Q:= OpenSQL(Req,True);
    TobDetail:= Tob.Create('MonDetail',Nil,-1);
    TobDetail.LoadDetailDB('MonDetail','','',Q,False);
    Ferme(Q);
    For x:=0 to TobLexique.Detail.Count-1 do
        begin
        Segment:= TobLexique.Detail[x].GetValue('PDL_DADSSEGMENT');
        Nature:= TobLexique.Detail[x].GetValue('PDL_DADSNATURE');
        OrdreSeg:= TobLexique.Detail[x].GetValue('ORDRESEG');
        Tablette:= TobLexique.Detail[x].GetValue('TABLETTE');
        If OrdreSeg<>'' then
           begin
           OrdreSegS:= OrdreSeg;
           OrdreSegI:= StrToInt(ReadTokenPipe(OrdreSegS,'L'));
           TD:= TobDetail.FindFirst(['PDS_SEGMENT','PDS_ORDRESEG'],
                                    [Segment,OrdreSegI],False);
           If TD<>Nil then
              begin
              ValeurSeg:= TD.GetValue('PDS_DONNEEAFFICH');
              If Tablette<>'' then
                 ValeurSeg:= RechDom(Tablette,ValeurSeg,False);
              end
           else
              ValeurSeg:= '';
           end
        else
           begin
           TD:= TobDetail.FindFirst(['PDS_SEGMENT'],[Segment],False);
           If TD<>nil then
              begin
              If TD.GetValue('PDS_DONNEEAFFICH')<>null then
                 ValeurSeg:= TD.GetValue('PDS_DONNEEAFFICH')
              else
                 ValeurSeg:= '';

              If (Tablette<>'') and (ValeurSeg<>'') then
                 ValeurSeg:= RechDom (Tablette,TD.GetValue('PDS_DONNEEAFFICH'),False);
                 If (Segment = 'S70.G01.00.010') and (TD.GetValue('PDS_DONNEEAFFICH')<>null) then ValeurSeg:= TD.GetValue('PDS_DONNEE'); //PT7
              end
           else
               if segment = 'S70.G01.00.015' then ValeurSeg := '31/12/'+AnneeCompta //PT7
              else ValeurSeg:= '';
           end;

        If OrdreSeg<>'' then
           Segment:= Segment+'.'+OrdreSeg;
        If Nature='N' then
           begin
           If IsNumeric (ValeurSeg) then
              T.AddchampSupValeur(Segment,Valeur(ValeurSeg),False)
           else
              T.AddchampSupValeur(Segment,0,False);
           end
        else
           T.AddchampSupValeur(Segment,ValeurSeg,False)
        end;
    FreeAndNil (TobDetail);
    MoveCurProgressForm ('Honoraire : '+
                         IntToStr(TobPeriodes.Detail[i].GetValue('PDE_ORDRE')));
    end;
FiniMoveProgressForm;
TobPeriodes.Detail.Sort('PDE_SALARIE;PDE_ORDRE');

FreeAndNil (TobLexique);

//PT8-2
CodeEtat:= 'PH'+AnneeF;
StModele:= 'SELECT MAX(MO_CODE) AS CODEETAT'+
           ' FROM MODELES WHERE'+
           ' MO_TYPE="E" AND'+
           ' MO_NATURE="PDA" AND'+
           ' MO_CODE LIKE "PH%"';
QModele:= OpenSQL(StModele, True);
if not QModele.eof then
   begin
   MaxModele:= QModele.FindField ('CODEETAT').AsString;
   if (CodeEtat>=MaxModele) then
      CodeEtat:= MaxModele;
   end;
Ferme (QModele);

TFQRS1(Ecran).CodeEtat:= CodeEtat;
//FIN PT8-2
TFQRS1(Ecran).LaTob:= TobPeriodes;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.OnLoad ;
begin
Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.OnArgument (S : String );
var
AnneeE, AnneePrec, ComboExer, MoisE, StPlus : string;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
begin
Inherited;
{PT8-1
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
SetControlText('ETAB1',Min);
SetControlText('ETAB2',Max);
}
SetControlText('DOSSIER', GetParamSoc ('SO_LIBELLE'));
THAnneeRef:=THValCombobox(GetControl('ANNEE'));
JourJ := Date;
DecodeDate(JourJ, AnneeA, MoisM, Jour);
if MoisM>9 then
   AnneePrec := IntToStr(AnneeA)
else
   AnneePrec := IntToStr(AnneeA-1);

if RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec) = TRUE then
   THAnneeRef.Value := ComboExer
else
   begin
   PGIBox('L''exercice '+AnneePrec+' n''existe pas', Ecran.Caption);
   end;

//PT9
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('ETAB', 'Plus', StPlus);
//FIN PT9
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... : 25/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.OnClose ;
begin
Inherited ;
FreeAndNil (TobPeriodes);	//PT6
end ;

//PT6
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/03/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : TOF;PGEDITDADSHON
*****************************************************************}
procedure TOF_PGEDITDADSHON.ChargerLexiqueDads (var TobLexique:tob;
                                                Annee:String);
var
Segment,Nature, StQ, DadsValeur : String;
TL : Tob;
Q : TQuery;
i,x, Position : Integer;
begin
StQ:= 'SELECT PDL_DADSSEGMENT, PDL_DADSNATURE, PDL_DADSVALEUR'+
      ' FROM DADSLEXIQUE WHERE'+
      ' PDL_DADSSEGMENT LIKE "S70%" AND'+
      ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
      ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
      ' ORDER BY PDL_DADSSEGMENT';
Q:= OpenSQL (StQ, True);
TobLexique.LoadDetailDB('LesSegments','','',Q,False);
Ferme(Q);

For i := 0 to TobLexique.Detail.Count - 1 do
    begin
    Segment := TobLexique.Detail[i].GetValue('PDL_DADSSEGMENT');
    Nature := TobLexique.Detail[i].GetValue('PDL_DADSNATURE');
    TobLexique.Detail[i].AddchampSupValeur('ORDRESEG','',False);
    TobLexique.Detail[i].AddchampSupValeur('TABLETTE','',False);
    If Segment = 'S70.G01.01.001' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','101');
       TobLexique.Detail[i].PutValue('TABLETTE','PGDADSHONTYPEREM');
       For x := 1 to 4 do
           begin
           TL := Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(101 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','PGDADSHONTYPEREM',False);
           end;
       end
    else
    If Segment = 'S70.G01.01.002' then
       begin
       TobLexique.Detail[i].PutValue('ORDRESEG','102');
       For x := 1 to 5 do
           begin
           TL := Tob.Create('UneFille',TobLexique,-1);
           TL.AddchampSupValeur('PDL_DADSSEGMENT',segment,False);
           TL.AddchampSupValeur('PDL_DADSNATURE',Nature,False);
           TL.AddchampSupValeur('ORDRESEG',IntToStr(102 + (x * 3)),False);
           TL.AddchampSupValeur('TABLETTE','',False);
           end;
       end
    else
    If (Segment='S70.G01.00.014') Then
       TobLexique.Detail[i].PutValue('TABLETTE','TTETABLISSEMENT')
    else
    if (TobLexique.Detail[i].GetValue('PDL_DADSVALEUR')<>'') then
       begin
       DadsValeur:= TobLexique.Detail[i].GetValue('PDL_DADSVALEUR');
       Position:= Pos ('T', DadsValeur);
       if (Position=1) then
          begin
          System.Delete (DadsValeur, 1, 2);
          Position:= Pos (';', DadsValeur);
          DadsValeur:= Copy (DadsValeur, 1, Position-1);
          TobLexique.Detail[i].PutValue('TABLETTE',DadsValeur);
          end;
       end;
    end;
end;
//FIN PT6

Initialization
  registerclasses ( [ TOF_PGEDITDADSHON ] ) ;
end.
