{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DADS2SALARIES (DADS2SALARIES)
Mots clefs ... : TOM;DADS2SALARIES
*****************************************************************}
{
PT1   : 20/01/2004 VG V_50 Adaptation cahier des charges 2003
PT2   : 07/10/2004 VG V_50 Contrôles de cohérence en validation - FQ N°11655
PT3   : 08/10/2004 VG V_50 Adaptation cahier des charges 2004
PT4   : 18/11/2004 VG V_60 Lors de la navigation, les coches n'étaient pas
                           décochées - FQ N°11789
PT5   : 28/12/2004 VG V_60 Contrôles de cohérence en validation - FQ N°11655
PT6   : 07/02/2005 VG V_60 Contrôles de cohérence en validation - FQ N°11971
PT7   : 16/03/2005 VG V_60 Initialisation des champs de l'enregistrement
                           "Salariés" en création - FQ N°12090
PT8   : 25/04/2005 VG V_60 La tablette des code emploi dépend d'un paramsoc                           
PT9   : 12/01/2006 VG V_65 "Conversion de type variant incorrecte" si on efface
                           le contenu de certaines zones numériques - FQ N°12099
PT10  : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Travail CIPDZ"
                           FQ N°14568
}
Unit UTOMDADS2SALARIES ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     Fiche,
{$ELSE}
     MaineAgl,
     eFiche,
     UTob,
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     UTOM,
     HTB97,
     PgOutils2,
     EntPaie,
     PgDADSCommun ;

Type
  TOM_DADS2SALARIES = Class (TOM)
  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  private
    PeriodeD1, PeriodeD2, PeriodeF1, PeriodeF2 : THEdit;
    AlloForfait, AuAvant, DepenAu, Employeur, Logement, Nourriture : TCheckBox;
    NTIC, Remboursement, Voiture : TCheckBox;
    BSalarie : TToolBarButton97;
    Annee, Salarie : string;
    Valid : TToolBarButton97;
    
    procedure SalarieClick (Sender: TObject);
    procedure Validation(Sender: TObject);
  end ;

Implementation

procedure TOM_DADS2SALARIES.OnNewRecord ;
begin
Inherited ;
//PT7
SetField('PD2_AVANTAGENATN', '    ');
SetField('PD2_NTIC', ' ');
SetField('PD2_FRAISPROFN', '    ');
//FIN PT7
end ;

procedure TOM_DADS2SALARIES.OnDeleteRecord ;
begin
Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2SALARIES.OnUpdateRecord ;
var
BufChamp, BufDest, BufOrig : string;
BufJourD, BufJourDChamp, BufJourF, BufJourFChamp : string;
BufMoisD, BufMoisDChamp, BufMoisF, BufMoisFChamp : string;
begin
Inherited ;
//PT9
if (GetControlText ('PD2_BASEBRUTE')='') then
   SetControlText ('PD2_BASEBRUTE', '0');
if (GetControlText ('PD2_AVANTAGENATV')='') then
   SetControlText ('PD2_AVANTAGENATV', '0');
if (GetControlText ('PD2_FRAISPROFM')='') then
   SetControlText ('PD2_FRAISPROFM', '0');
if (GetControlText ('PD2_REVENUSACTIV')='') then
   SetControlText ('PD2_REVENUSACTIV', '0');
if (GetControlText ('PD2_INDEMEXPATRI')='') then
   SetControlText ('PD2_INDEMEXPATRI', '0');
if (GetControlText ('PD2_NBREHEURE')='') then
   SetControlText ('PD2_NBREHEURE', '0');
//FIN PT9

BufOrig:= GetField ('PD2_SIRET');
ForceNumerique (BufOrig, BufDest);
if ControlSiret (BufDest)=False then
   begin
   LastError:= 1;
   PgiBox ('Le SIRET de l''établissement#13#10'+
           '"'+GetField ('PD2_ETABLISSEMENT')+'" n''est pas valide',
           Ecran.Caption);
   SetFocusControl ('PD2_ETABLISSEMENT');
   exit;
   end;

BufChamp:= GetField ('PD2_PAYS');
if (BufChamp='FRA') then
   begin
   if ((GetField ('PD2_COMMUNENAISS')='') or
      (GetField ('PD2_COMMUNENAISS')=null)) then
      begin
      LastError:= 1;
      PgiBox ('La commune de naissance n''est pas renseignée', Ecran.Caption);
      exit;
      end;

   if ((GetField ('PD2_DEPART')='') or (GetField ('PD2_DEPART')=null)) then
      begin
      LastError:= 1;
      PgiBox ('Le département de naissance n''est pas renseigné',
              Ecran.Caption);
      exit;
      end;
   end
else
   if ((BufChamp='') or (BufChamp=null)) then
      begin
      LastError:= 1;
      PgiBox ('Le pays de naissance n''est pas renseigné', Ecran.Caption);
      exit;
      end;

if ((GetField ('PD2_CIVILITE')='') or (GetField ('PD2_CIVILITE')=null)) then
   begin
   LastError:= 1;
   PgiBox ('La civilité n''est pas renseignée', Ecran.Caption);
   exit;
   end;

if ((GetField ('PD2_PRENOMDADS')='') or (GetField ('PD2_PRENOMDADS')=null)) then
   begin
   LastError:= 1;
   PgiBox ('Le prénom n''est pas renseignée', Ecran.Caption);
   exit;
   end;

if ((GetField ('PD2_CODEPOSTAL')='') or (GetField ('PD2_CODEPOSTAL')=null) or
   (GetField ('PD2_CODEPOSTAL')<'00000') or
   (GetField ('PD2_CODEPOSTAL')>'99999')) then
   begin
   LastError:= 1;
   PgiBox ('Le code postal est incorrect', Ecran.Caption);
   SetFocusControl ('PD2_CODEPOSTAL');
   exit;
   end;

if ((GetField ('PD2_BUREAUDISTRIB')='') or
   (GetField ('PD2_BUREAUDISTRIB')=null)) then
   begin
   LastError:= 1;
   PgiBox ('Le bureau distributeur n''est pas renseigné', Ecran.Caption);
   SetFocusControl ('PD2_BUREAUDISTRIB');
   exit;
   end;

if ((GetField ('PD2_EMPLOIQUALIF')='') or
   (GetField ('PD2_EMPLOIQUALIF')=null)) then
   begin
   LastError:= 1;
   PgiBox ('La nature de l''emploi n''est pas renseignée', Ecran.Caption);
   SetFocusControl ('PD2_EMPLOIQUALIF');
   exit;
   end;

if ((GetField ('PD2_CONDEMPLOI')<>'C') and
   (GetField ('PD2_CONDEMPLOI')<> 'I') and
   (GetField ('PD2_CONDEMPLOI')<> 'P') and
   (GetField ('PD2_CONDEMPLOI')<> 'D')) then
   begin
   LastError:= 1;
{PT10
   PgiBox ('La condition d''emploi n''est pas correctement renseignée',
           Ecran.Caption);
}
   PgiBox ('Le travail CIPDZ n''est pas correctement renseigné',
           Ecran.Caption);
//FIN PT10
   SetFocusControl ('PD2_CONDEMPLOI');
   exit;
   end;

if (GetField ('PD2_BASEBRUTE')=0) then
   begin
   PgiBox ('Attention ! La base brute fiscale est égale à 0', Ecran.Caption);
   SetFocusControl ('PD2_BASEBRUTE');
   end;

if ((PeriodeD1<>nil) and (PeriodeF1<>nil)) then
   begin
   BufJourD:= Copy (PeriodeD1.Text, 1, 2);
   BufMoisD:= Copy (PeriodeD1.Text, 4, 2);
   BufJourF:= Copy (PeriodeF1.Text, 1, 2);
   BufMoisF:= Copy (PeriodeF1.Text, 4, 2);
   BufChamp:= GetField ('PD2_PERIODE1');
   BufJourDChamp:= Copy (BufChamp, 1, 2);
   BufMoisDChamp:= Copy (BufChamp, 3, 2);
   BufJourFChamp:= Copy (BufChamp, 5, 2);
   BufMoisFChamp:= Copy (BufChamp, 7, 2);
   if ((BufJourD<>BufJourDChamp) or (BufMoisD<>BufMoisDChamp) or
      (BufJourF<>BufJourFChamp) or (BufMoisF<>BufMoisFChamp)) then
      SetField ('PD2_PERIODE1', BufJourD+BufMoisD+BufJourF+BufMoisF);
   end;

if ((PeriodeD2<>nil) and (PeriodeF2<>nil)) then
   begin
   BufJourD:= Copy (PeriodeD2.Text, 1, 2);
   BufMoisD:= Copy (PeriodeD2.Text, 4, 2);
   BufJourF:= Copy (PeriodeF2.Text, 1, 2);
   BufMoisF:= Copy (PeriodeF2.Text, 4, 2);
   BufChamp:= GetField ('PD2_PERIODE2');
   BufJourDChamp:= Copy (BufChamp, 1, 2);
   BufMoisDChamp:= Copy (BufChamp, 3, 2);
   BufJourFChamp:= Copy (BufChamp, 5, 2);
   BufMoisFChamp:= Copy (BufChamp, 7, 2);
   if ((BufJourD<>BufJourDChamp) or (BufMoisD<>BufMoisDChamp) or
      (BufJourF<>BufJourFChamp) or (BufMoisF<>BufMoisFChamp)) then
      SetField ('PD2_PERIODE2', BufJourD+BufMoisD+BufJourF+BufMoisF);
   end;

//PT6
if ((Nourriture.Checked=False) and (Logement.Checked=False) and
   (Voiture.Checked=False) and (AuAvant.Checked=False) and
   (NTIC.Checked=False) and (GetField ('PD2_AVANTAGENATV')<>0)) then
   begin
   LastError:= 1;
   PgiBox ('Le montant des avantages en nature est égal à '+
           FloatToStr(GetField('PD2_AVANTAGENATV'))+'#13#10'+
           'mais aucune zone correspondante n''est cochée', Ecran.Caption);
   SetFocusControl ('PD2_AVANTAGENATV');
   exit;
   end;
//FIN PT6

BufChamp:= GetField ('PD2_AVANTAGENATN');
if ((Nourriture<>nil) and (Logement<>nil) and (Voiture<>nil) and
   (AuAvant<>nil)) then
   if (((Copy (BufChamp, 1, 1)='N')<>Nourriture.Checked) or
      ((Copy (BufChamp,2,1)='L')<>Logement.Checked) or
      ((Copy (BufChamp,3,1)='V')<>Voiture.Checked) or
      ((Copy (BufChamp,4,1)='A')<>AuAvant.Checked)) then
      begin
      if (Nourriture.Checked=TRUE) then
         BufChamp:= 'N'
      else
         BufChamp:= ' ';

      if (Logement.Checked=TRUE) then
         BufChamp:= BufChamp+'L'
      else
         BufChamp:= BufChamp+' ';

      if (Voiture.Checked=TRUE) then
         BufChamp:= BufChamp+'V'
      else
         BufChamp:= BufChamp+' ';

      if (AuAvant.Checked=TRUE) then
         BufChamp:= BufChamp+'A'
      else
         BufChamp:= BufChamp+' ';
      SetField ('PD2_AVANTAGENATN', BufChamp);
      end;

BufChamp:= GetField ('PD2_NTIC');
if (NTIC<>nil) then
   if ((BufChamp='T')<>NTIC.Checked) then
      begin
      if (NTIC.Checked=TRUE) then
         BufChamp:= 'T'
      else
         BufChamp:= ' ';
      SetField ('PD2_NTIC', BufChamp);
      end;

if ((AlloForfait.Checked=False) and (Remboursement.Checked=False) and
   (Employeur.Checked=False) and (DepenAu.Checked=False) and
   (GetField ('PD2_FRAISPROFM')<>0)) then
   begin
   LastError:= 1;
   PgiBox ('Le montant des frais professionnels est égal à '+
           FloatToStr(GetField('PD2_FRAISPROFM'))+'#13#10'+
           'mais aucune zone correspondante n''est cochée', Ecran.Caption);
   SetFocusControl ('PD2_FRAISPROFM');
   exit;
   end;

BufChamp:= GetField ('PD2_FRAISPROFN');
if ((AlloForfait<>nil) and (Remboursement<>nil) and (Employeur<>nil) and
   (DepenAu<>nil)) then
   if (((Copy (BufChamp, 1, 1)='F')<>AlloForfait.Checked) or
      ((Copy (BufChamp, 2, 1)='R')<>Remboursement.Checked) or
      ((Copy (BufChamp, 3, 1)='P')<>Employeur.Checked) or
      ((Copy (BufChamp, 4, 1)='A')<>DepenAu.Checked)) then
      begin
      if (AlloForfait.Checked=TRUE) then
         BufChamp:= 'F'
      else
         BufChamp:= ' ';

      if (Remboursement.Checked=TRUE) then
         BufChamp:= BufChamp+'R'
      else
         BufChamp:= BufChamp+' ';

      if (Employeur.Checked=TRUE) then
         BufChamp:= BufChamp+'P'
      else
         BufChamp:= BufChamp+' ';

      if (DepenAu.Checked=TRUE) then
         BufChamp:= BufChamp+'A'
      else
         BufChamp:= BufChamp+' ';

      SetField ('PD2_FRAISPROFN', BufChamp);
      end;

if ((GetField ('PD2_REVENUSACTIV')=0) and (GetField ('PD2_INDEMEXPATRI')=0) and
   (GetField ('PD2_AVANTAGENATV')=0)) then
   begin
   if (GetField ('PD2_BASEBRUTE')=0) then
      SetField ('PD2_CODEABSENCE', 'I')
   else
      SetField ('PD2_CODEABSENCE', 'K');
   end
else
   SetField ('PD2_CODEABSENCE', ' ');

if (GetField ('PD2_NBREHEURE')<0) then
   begin
   LastError:= 1;
   PgiBox ('Le nombre d''heures rémunérées est incorrect', Ecran.Caption);
   SetFocusControl ('PD2_NBREHEURE');
   exit;
   end;
end ;

procedure TOM_DADS2SALARIES.OnAfterUpdateRecord ;
begin
Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2SALARIES.OnLoadRecord ;
var
Buf, BufAn, BufJour, BufMois : string;
AnneeA, Jour, MoisM : Word;
begin
Inherited ;
DecodeDate(DebExer, AnneeA, MoisM, Jour);
BufJour:='';
BufMois:='';
BufAn:='';
if ((PeriodeD1 <> nil) and (PeriodeF1 <> nil)) then
   begin
   BufJour := Copy(GetField('PD2_PERIODE1'), 1, 2);
   BufMois := Copy(GetField('PD2_PERIODE1'), 3, 2);
   if ((BufMois <> '') and (BufMois <> '  ')) then
      begin
      if (StrToInt(BufMois) < MoisM) then
         BufAn := IntToStr(AnneeA+1)
      else
         BufAn := IntToStr(AnneeA);
      end
   else
      BufAn:='';
   PeriodeD1.Text := BufJour+'/'+BufMois+'/'+BufAn;

   BufJour := Copy(GetField('PD2_PERIODE1'), 5, 2);
   BufMois := Copy(GetField('PD2_PERIODE1'), 7, 2);
   if ((BufMois <> '') and (BufMois <> '  ')) then
      begin
      if (StrToInt(BufMois) < MoisM) then
         BufAn := IntToStr(AnneeA+1)
      else
         BufAn := IntToStr(AnneeA);
      end
   else
      BufAn:='';
   PeriodeF1.Text := BufJour+'/'+BufMois+'/'+BufAn;
   end;

BufJour:='';
BufMois:='';
BufAn:='';
if ((PeriodeD2 <> nil) and (PeriodeF2 <> nil)) then
   begin
   BufJour := Copy(GetField('PD2_PERIODE2'), 1, 2);
   BufMois := Copy(GetField('PD2_PERIODE2'), 3, 2);
   if ((BufMois <> '') and (BufMois <> '  ')) then
      begin
      if (StrToInt(BufMois) < MoisM) then
         BufAn := IntToStr(AnneeA+1)
      else
         BufAn := IntToStr(AnneeA);
      end
   else
      BufAn:='';
   PeriodeD2.Text := BufJour+'/'+BufMois+'/'+BufAn;

   BufJour := Copy(GetField('PD2_PERIODE2'), 5, 2);
   BufMois := Copy(GetField('PD2_PERIODE2'), 7, 2);
   if ((BufMois <> '') and (BufMois <> '  ')) then
      begin
      if (StrToInt(BufMois) < MoisM) then
         BufAn := IntToStr(AnneeA+1)
      else
         BufAn := IntToStr(AnneeA);
      end
   else
      BufAn:='';
   PeriodeF2.Text := BufJour+'/'+BufMois+'/'+BufAn;
   end;

Buf := GetField('PD2_AVANTAGENATN');
{PT4
if ((Copy(Buf, 1, 1) = 'N') and (Nourriture <> nil)) then
   Nourriture.Checked := TRUE;
if ((Copy(Buf, 2, 1) = 'L') and (Logement <> nil)) then
   Logement.Checked := TRUE;
if ((Copy(Buf, 3, 1) = 'V') and (Voiture <> nil)) then
   Voiture.Checked := TRUE;
if ((Copy(Buf, 4, 1) = 'A') and (AuAvant <> nil)) then
   AuAvant.Checked := TRUE;
}
if (Nourriture <> nil) then
   begin
   if (Copy(Buf, 1, 1)='N') then
      Nourriture.Checked:= TRUE
   else
      Nourriture.Checked:= FALSE;
   end;
if (Logement <> nil) then
   begin
   if (Copy(Buf, 2, 1)='L') then
      Logement.Checked:= TRUE
   else
      Logement.Checked:= FALSE;
   end;
if (Voiture <> nil) then
   begin
   if (Copy(Buf, 3, 1)='V') then
      Voiture.Checked:= TRUE
   else
      Voiture.Checked:= FALSE;
   end;
if (AuAvant <> nil) then
   begin
   if (Copy(Buf, 4, 1)='A') then
      AuAvant.Checked:= TRUE
   else
      AuAvant.Checked:= FALSE;
   end;
//FIN PT4

//PT3
{PT4
Buf := GetField('PD2_NTIC');
if ((Buf='T') and (NTIC <> nil)) then
   NTIC.Checked := TRUE;
}
Buf := GetField('PD2_NTIC');
if (NTIC <> nil) then
   begin
   if (Buf='T') then
      NTIC.Checked:= TRUE
   else
      NTIC.Checked:= FALSE;
   end;
//FIN PT4
//FIN PT3

Buf := GetField('PD2_FRAISPROFN');
{PT4
if ((Copy(Buf, 1, 1) = 'F') and (AlloForfait <> nil)) then
   AlloForfait.Checked := TRUE;
if ((Copy(Buf, 2, 1) = 'R') and (Remboursement <> nil)) then
   Remboursement.Checked := TRUE;
if ((Copy(Buf, 3, 1) = 'P') and (Employeur <> nil)) then
   Employeur.Checked := TRUE;
if ((Copy(Buf, 4, 1) = 'A') and (DepenAu <> nil)) then
   DepenAu.Checked := TRUE;
}
if (AlloForfait <> nil) then
   begin
   if (Copy(Buf, 1, 1)='F') then
      AlloForfait.Checked:= TRUE
   else
      AlloForfait.Checked:= FALSE;
   end;
if (Remboursement <> nil) then
   begin
   if (Copy(Buf, 2, 1)='R') then
      Remboursement.Checked:= TRUE
   else
      Remboursement.Checked:= FALSE;
   end;
if (Employeur <> nil) then
   begin
   if (Copy(Buf, 3, 1)='P') then
      Employeur.Checked:= TRUE
   else
      Employeur.Checked:= FALSE;
   end;
if (DepenAu <> nil) then
   begin
   if (Copy(Buf, 4, 1)='A') then
      DepenAu.Checked:= TRUE
   else
      DepenAu.Checked:= FALSE;
   end;
//FIN PT4

{PT1
if ((GetField('PD2_REMNETTE') = 0) and (GetField('PD2_INDEMEXPATRI') = 0) and
}
if ((GetField('PD2_REVENUSACTIV') = 0) and
   (GetField('PD2_INDEMEXPATRI') = 0) and
//FIN PT1
   (GetField('PD2_AVANTAGENATV') = 0)) then
   begin
   if (GetField('PD2_BASEBRUTE') = 0) then
      SetField('PD2_CODEABSENCE', 'I')
   else
      SetField('PD2_CODEABSENCE', 'K');
   end
else
   SetField('PD2_CODEABSENCE', ' ');
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2SALARIES.OnChangeField ( F: TField );
var
StEtab : string;
QRechEtab : TQuery;
begin
Inherited;
//PT2
if (F.FieldName = 'PD2_ETABLISSEMENT') then
   begin
   StEtab:= 'SELECT ET_ETABLISSEMENT, ET_SIRET, ETB_DADSSECTION,'+
            ' ETB_TYPDADSSECT'+
            ' FROM ETABLISS'+
            ' LEFT JOIN ETABCOMPL ON'+
            ' ETB_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
            ' ET_ETABLISSEMENT="'+GetField('PD2_ETABLISSEMENT')+'"';
   QRechEtab:=OpenSql(StEtab, TRUE);
   if (not QRechEtab.EOF) then
      begin
      SetField ('PD2_SIRET', QRechEtab.FindField('ET_SIRET').AsString);
      SetField ('PD2_SECTIONETAB',
                QRechEtab.FindField('ETB_DADSSECTION').AsString);
      SetField ('PD2_TYPEDADS',
                QRechEtab.FindField('ETB_TYPDADSSECT').AsString);
      end;
   Ferme(QRechEtab);
   end;
//FIN PT2
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2SALARIES.OnArgument ( S: String ) ;
var
Pages : TPageControl;
Buf : string;
begin
Inherited ;
Buf:=Trim(ReadTokenPipe(S,';')) ;
Salarie:=Trim(ReadTokenPipe(S,';')) ;
Annee:=Trim(ReadTokenPipe(S,';')) ;

//PT8
if VH_Paie.PGPCS2003 then
   SetControlProperty('PD2_CODEEMPLOI', 'datatype', 'PGCODEPCSESE');
//FIN PT8

PeriodeD1 := THEdit (GetControl ('EPERIODE1D'));
{PT5
if (PeriodeD1 <> nil) then
   PeriodeD1.OnExit:=CExit;
}
PeriodeF1 := THEdit (GetControl ('EPERIODE1F'));
{PT5
if (PeriodeF1 <> nil) then
   PeriodeF1.OnExit:=CExit;
}
PeriodeD2 := THEdit (GetControl ('EPERIODE2D'));
{PT5
if (PeriodeD2 <> nil) then
   PeriodeD2.OnExit:=CExit;
}
PeriodeF2 := THEdit (GetControl ('EPERIODE2F'));
{PT5
if (PeriodeF2 <> nil) then
   PeriodeF2.OnExit:=CExit;
}

Nourriture := TCheckBox (GetControl ('CHNOURRITURE'));
{PT5
if (Nourriture <> nil) then
   Nourriture.OnExit:=CExit;
}
Logement := TCheckBox (GetControl ('CHLOGEMENT'));
{PT5
if (Logement <> nil) then
   Logement.OnExit:=CExit;
}
Voiture := TCheckBox (GetControl ('CHVOITURE'));
{PT5
if (Voiture <> nil) then
   Voiture.OnExit:=CExit;
}
AuAvant := TCheckBox (GetControl ('CHAUTRESAVANTAGES'));
{PT5
if (AuAvant <> nil) then
   AuAvant.OnExit:=CExit;
}
NTIC := TCheckBox (GetControl ('CHNTIC'));
{PT5
if (NTIC <> nil) then
   NTIC.OnExit:=CExit;
}
AlloForfait := TCheckBox (GetControl ('CHALLOCFORFAIT'));
{PT5
if (AlloForfait <> nil) then
   AlloForfait.OnExit:=CExit;
}
Remboursement := TCheckBox (GetControl ('CHREMBOURSEMENTS'));
{PT5
if (Remboursement <> nil) then
   Remboursement.OnExit:=CExit;
}
Employeur := TCheckBox (GetControl ('CHEMPLOYEUR'));
{PT5
if (Employeur <> nil) then
   Employeur.OnExit:=CExit;
}
DepenAu := TCheckBox (GetControl ('CHDEPENSES'));
{PT5
if (DepenAu <> nil) then
   DepenAu.OnExit:=CExit;
}

// Positionnement sur le premier onglet
Pages := TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;

BSalarie := TToolbarButton97(GetControl('BSALARIE'));
if BSalarie <> NIL then
   BSalarie.OnClick := SalarieClick;

//PT5
// Gestion du navigateur
Valid := TToolBarButton97(GetControl('BValider'));
if Valid<>nil then
   Valid.OnClick:=Validation;
//FIN PT5
end ;

procedure TOM_DADS2SALARIES.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_DADS2SALARIES.OnCancelRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Click sur le bouton Salarié
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2SALARIES.SalarieClick(Sender: TObject);
var
Sal : string;
begin
Sal := GetField('PD2_SALARIE');
AGLLanceFiche ('PAY', 'SALARIE', '',Sal, '');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Click sur une donnée autre que DB
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
{PT5
procedure TOM_DADS2SALARIES.CExit(Sender: TObject);
begin
if not (ds.state in [dsinsert,dsedit]) then ds.edit;
end;
}

//PT5
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/12/2004
Modifié le ... :   /  /
Description .. : Validation
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOM_DADS2SALARIES.Validation(Sender: TObject);
begin
if not (ds.state in [dsinsert,dsedit]) then
   ds.edit;
TFFiche(Ecran).BValiderClick(Nil);
end;
//FIN PT5

Initialization
  registerclasses ( [ TOM_DADS2SALARIES ] ) ;
end.
