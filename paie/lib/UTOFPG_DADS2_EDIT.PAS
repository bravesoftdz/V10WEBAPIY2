{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/08/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DADSB_EDITION ()
Mots clefs ... : TOF;PGEDIT_DADSB
*****************************************************************}
{
PT1   : 24/09/2004 VG V_50 Pour l'édition des honoraires, on ne sélectionne pas
                           les codes salariés
PT2   : 11/10/2004 VG V_50 Liste d'exportation non disponible - FQ N°11659
PT3   : 26/01/2006 VG V_65 Saut de page si 2 salariés n'appartiennent pas au
                           même établissement - FQ N°11659
PT4   : 06/04/2006 VG V_65 Edition des traitements et salaires - FQ N°12151
PT5   : 02/02/2007 V_80 FC Mise en place filtrage des habilitations/poupulations
}
Unit UTOFPG_DADS2_EDIT ;

Interface

Uses StdCtrls,
     Classes,
{$IFNDEF EAGLCLIENT}
     QRS1,
     FE_Main,
{$ELSE}
     eQRS1,Utob,
     MaineAGL,
{$ENDIF}
     sysutils,
     HCtrls,
     HEnt1,
     UTOF,
//     PGEditOutils2,
{$IFNDEF COMPTA}
     PgOutils2,
{$ENDIF}
     EntPaie,
     ParamSoc
{$IFNDEF COMPTA}
{$IFNDEF DADSB}
     , P5Def //PT5
{$ENDIF}
{$ENDIF}
      ;

Type
  TOF_PGEDIT_DADSB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    Honor, Salaires : TCheckBox;
    Function ConstruireRequete(RuptEtab,OrdreAlpha,Honor:Boolean):String;
    procedure HonorChange(Sender: TObject);
    procedure SalairesChange(Sender: TObject);
  end ;

  procedure LanceFiche_DADS2Edition;

Implementation

procedure LanceFiche_DADS2Edition;
begin
  AGLLanceFiche('PAY', 'DADSB_EDITION', '', '', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnUpdate ;
var
Annee, CDate, CEtab, CSal1, Etab, MCEtab, MCSal, OrderBy, PGExercice : String;
{$IFDEF COMPTA}
  CetabIn : String;
{$ENDIF}
requete, Salarie, SQL : String;
TriSal:TCheckBox;
Habilitation:String;//PT5
begin
Inherited ;
OrderBy := '';
TriSal := TCheckBox(GetControl('TRISAL'));
{PT4
Etab1 := GetControlText('ETAB1');
Etab2 := GetControlText('ETAB2');
PGGlbEtabDe := Etab1;
PGGlbEtabA := Etab2;
Sal1 := GetControlText('SAL1');
Sal2 := GetControlText('SAL2');
PGGlbSalDe := Sal1;
PGGlbSalA := Sal2;
}
Annee := GetControlText('ANNEE');

PGExercice:= Rechdom ('PGANNEE', Annee, False);

if (Honor.checked = True) then
   begin
   CDate := ' WHERE PDH_VALIDITE ="'+PGExercice+'"';
{PT4
   CEtab:= ' AND ET_ETABLISSEMENT>="'+Etab1+'" AND'+
           ' ET_ETABLISSEMENT<="'+Etab2+'"';
}

{$IFDEF COMPTA}
  if THMultiValComboBox(Getcontrol('ETAB')).Tous then
    CEtab               := ''
  else
  begin
    CEtab               := ' AND PDH_ETABLISSEMENT IN (';
    MCEtab              := GetControlText('ETAB');
    Etab                := READTOKENST(MCEtab);
    CetabIn             := '';
    While Etab <> '' do
    begin
      if CetabIn <> '' then
        CetabIn         := CetabIn + ', ';
      CetabIn           := CetabIn + '"' + Etab + '"';
      Etab              := READTOKENST(MCEtab);
    end;
    CEtab               := CEtab + CetabIn + ') ';
  end;
{$ENDIF}
   If (GetControlText('CRUPTETAB')='X') then
      begin
      If (TriSal.Checked = True) Then
         begin
         Requete := ConstruireRequete(True, True, True);
         OrderBy := ' ORDER BY ET_ETABLISSEMENT,PDH_NOMBEN'+
                    ' GROUP BY ET_ETABLISSEMENT';
         end
      Else
         begin
         Requete := ConstruireRequete(True, False, True);
         OrderBy := ' ORDER BY ET_ETABLISSEMENT, PDH_HONORAIRE';
         end;
      end
   Else
      begin
{$IFNDEF COMPTA}
      CEtab := '';
{$ENDIF}
      If (TriSal.Checked = True) Then
         begin
         Requete := ConstruireRequete(False, True, True);
         OrderBy := ' ORDER BY PDH_NOMBEN';
         end
      Else
         begin
         Requete := ConstruireRequete(False, False, True);
         OrderBy := ' ORDER BY PDH_HONORAIRE';
         end;
      end;
   end
Else
   begin
   CDate:= ' WHERE PD2_VALIDITE ="'+PGExercice+'"';
{PT4
   If (Sal1 <> '') Then
      CSal1 := ' AND PD2_SALARIE>="'+Sal1+'"';
   If (Sal2 <> '') Then
      CSal2 := ' AND PD2_SALARIE<="'+Sal2+'"';
}
   if (THMultiValCombobox(GetControl('SAL')).Tous) then
      CSal1:= ''
   else
      begin
      MCSal:= GetControlText('SAL');
      Salarie:= ReadTokenpipe(MCSal, ';');
      CSal1:= ' AND (';
      While (Salarie <> '') do
            begin
            CSal1:= CSal1+' PD2_SALARIE="'+Salarie+'"';
            Salarie := ReadTokenpipe(MCSal,';');
            if (Salarie <> '') then
               CSal1:= CSal1+' OR'
            else
               CSal1:= CSal1+')';
            end;
      end;
//FIN PT4


{PT4
   CEtab:= 'AND PD2_ETABLISSEMENT>="'+Etab1+'" AND'+
           ' PD2_ETABLISSEMENT<="'+Etab2+'"';
}
   if (THMultiValCombobox(GetControl('ETAB')).Tous) then
      CEtab:= ''
   else
      begin
      MCEtab:= GetControlText('ETAB');
      Etab:= ReadTokenpipe(MCEtab, ';');
      CEtab:= ' AND (';
      While (Etab <> '') do
            begin
            CEtab:= CEtab+' PD2_ETABLISSEMENT="'+Etab+'"';
            Etab := ReadTokenpipe(MCEtab,';');
            if (Etab <> '') then
               CEtab:= CEtab+' OR'
            else
               CEtab:= CEtab+')';
            end;
      end;
//FIN PT4

   If (GetControlText('CRUPTETAB')='X') then
      begin
      SetControlText ('XX_RUPTURE', 'PD2_ETABLISSEMENT');                  //PT3
      If (TriSal.Checked = True) Then
         begin
         Requete := ConstruireRequete(True, True, False);
         OrderBy := ' ORDER BY ET_ETABLISSEMENT,PSA_LIBELLE';
         end
      Else
         begin
         Requete := ConstruireRequete(True, False, False);
         OrderBy := ' ORDER BY ET_ETABLISSEMENT, PD2_SALARIE';
         end;
      end
   Else
      begin
      SetControlText ('XX_RUPTURE', '');         //PT3
      If (TriSal.Checked = True) Then
         begin
{PT4
         Requete := ConstruireRequete(False, True, False);
}
         If (GetControlText('SALAIRES')='-') then
            Requete := ConstruireRequete(False, True, False)
         else
            Requete := ConstruireRequete(True, True, False);
//FIN PT4
         OrderBy := ' ORDER BY PSA_LIBELLE';
         end
      Else
         begin
{PT4
         Requete := ConstruireRequete(False, False, False);
}
         If (GetControlText('SALAIRES')='-') then
            Requete := ConstruireRequete(False, False, False)
         else
            Requete := ConstruireRequete(True, False, False);
//FIN PT4
         OrderBy := ' ORDER BY PD2_SALARIE';
         end;
      end;
   end;

If Honor.Checked=True Then
   TFQRS1(Ecran).CodeEtat:='PDH'
Else
{PT4
   TFQRS1(Ecran).CodeEtat:= 'PDC';
}
if (GetControlText('SALAIRES')='-') then
   TFQRS1(Ecran).CodeEtat:= 'PDC'
else
   TFQRS1(Ecran).CodeEtat:= 'PDT';

{PT4
SQL:=requete + CDate + CSal1 + CSal2 + CEtab + OrderBy;
}
  //DEB PT5
  Habilitation := '';
{$IFNDEF COMPTA}
{$IFNDEF DADSB}
  if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
  begin
    if (CDate<>'') or (CSal1<>'') or (CEtab<>'') then
      Habilitation := ' AND PD2_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')'
    else
      Habilitation := ' WHERE PD2_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + ')';
  end;
{$ENDIF}
{$ENDIF}
  //FIN PT5

SQL:= requete + CDate + CSal1 + CEtab +Habilitation+OrderBy; //PT5
//FIN PT4
TFQRS1(Ecran).WhereSQL := SQL;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnArgument (S : String ) ;
var
AnneePrec : String;
AnneeA, JourJ, MoisM : Word;
Jour : TDateTime;
begin
Inherited;
{PT4
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
SetControlText('ETAB1',Min);
SetControlText('ETAB2',Max);
RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
SetControlText('Sal1',Min);
SetControlText('Sal2',Max);
}
SetControlText('DOSSIER', GetParamSoc ('SO_LIBELLE'));

{PT4
THSal1:=ThEdit(getcontrol('SAL1'));
If THSal1<>nil then
   THSal1.OnExit:=ExitEdit;
THSal2:=ThEdit(getcontrol('SAL2'));
If THSal2<>nil then
   THSal2.OnExit:=ExitEdit;
}

Honor:= TCheckBox(GetControl('HONOR'));
if (Honor<>nil) then
   Honor.OnClick:= HonorChange;

//PT4
Salaires:= TCheckBox(GetControl('SALAIRES'));
if (Salaires<>nil) then
   Salaires.OnClick:= SalairesChange;
//FIN PT4

{$IFDEF COMPTA}
  //on ne gère que les honoraires
  Honor.Checked         := True;
  Honor.Visible         := False;
  Salaires.Visible      := False;
{$ENDIF}

Jour := Date;
DecodeDate(Jour, AnneeA, MoisM, JourJ);
if MoisM>9 then
   AnneePrec := IntToStr(AnneeA)
else
   AnneePrec := IntToStr(AnneeA-1);

SetControlText('ANNEE', copy(AnneePrec, 1, 1)+copy(AnneePrec, 3, 2));
SetControlVisible('FListe', False);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnClose ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDIT_DADSB.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/08/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
Function TOF_PGEDIT_DADSB.ConstruireRequete(RuptEtab,OrdreAlpha,Honor:Boolean):String;
var Requete:String;
begin
If (Honor = True) then
   begin
   If (RuptEtab=True) then
      begin
      requete:='SELECT D1.*, ET_ETABLISSEMENT'+
               ' FROM ETABLISS'+
               ' LEFT JOIN DADS2HONORAIRES D1 ON'+
               ' D1.PDH_ETABLISSEMENT=ET_ETABLISSEMENT';
      end
   Else
      begin
      requete:='SELECT *'+
               ' FROM DADS2HONORAIRES';
      end;
   end
else
   begin
   If (RuptEtab=True) then
      begin
      If (OrdreAlpha=True) then
         begin
         requete:= 'SELECT D1.*, ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1,'+
                   ' ET_ADRESSE2, ET_CODEPOSTAL, ET_VILLE, ET_SIRET, ET_APE,'+
                   ' PSA_LIBELLE'+
                   ' FROM ETABLISS'+
                   ' LEFT JOIN DADS2SALARIES D1 ON'+
                   ' D1.PD2_ETABLISSEMENT=ET_ETABLISSEMENT'+
                   ' LEFT JOIN SALARIES ON'+
                   ' D1.PD2_SALARIE=PSA_SALARIE';
         end
      Else
         begin
         requete:= 'SELECT D1.*, ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1,'+
                   ' ET_ADRESSE2, ET_CODEPOSTAL, ET_VILLE, ET_SIRET, ET_APE'+
                   ' FROM ETABLISS'+
                   ' LEFT JOIN DADS2SALARIES D1 ON'+
                   ' D1.PD2_ETABLISSEMENT=ET_ETABLISSEMENT';
         end;
      end
   Else
      begin
      If (OrdreAlpha=True) then
         begin
         requete:='SELECT D1.*, PSA_LIBELLE'+
                  ' FROM DADS2SALARIES D1'+
                  ' LEFT JOIN SALARIES ON'+
                  ' D1.PD2_SALARIE=PSA_SALARIE';
         end
      Else
         begin
         requete:='SELECT *'+
                  ' FROM DADS2SALARIES';
         end;
      end;
   end;

Result:=Requete;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/09/2004
Modifié le ... :   /  /
Description .. : Clic sur la checkbox "Honoraires"
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PGEDIT_DADSB.HonorChange(Sender: TObject);
begin
if Honor <> NIL then
   begin
   if Honor.Checked = FALSE then
      begin
      SetControlVisible ('TSAL1', TRUE);
{PT4
      SetControlVisible ('SAL1', TRUE);
      SetControlVisible ('TSAL2', TRUE);
      SetControlVisible ('SAL2', TRUE);
}
      SetControlVisible ('SAL', TRUE);
//FIN PT4
      SetControlVisible ('TRISAL', TRUE);
{PT4
      RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
      SetControlText('SAL1',Min);
      SetControlText('SAL2',Max);
}
      end
   else
      begin
      SetControlVisible ('TSAL1', FALSE);
{PT4
      SetControlVisible ('SAL1', FALSE);
      SetControlVisible ('TSAL2', FALSE);
      SetControlVisible ('SAL2', FALSE);
}
      SetControlVisible ('SAL', FALSE);
//FIN PT4
      SetControlVisible ('TRISAL', FALSE);
{PT4
      SetControlText ('SAL1','');
      SetControlText ('SAL2','');
}
      SetControlText ('SAL','');
//FIN PT4
      SetControlText ('TRISAL','-');
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/04/2006
Modifié le ... :   /  /
Description .. : Clic sur la checkbox "Traitements et salaires"
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PGEDIT_DADSB.SalairesChange(Sender: TObject);
begin
if Salaires <> NIL then
   begin
   if Salaires.Checked = FALSE then
      SetControlEnabled ('HONOR', TRUE)
   else
      begin
      SetControlEnabled ('HONOR', FALSE);
      SetControlText ('HONOR','-');
      end;
   end;
end;

Initialization
  registerclasses ( [ TOF_PGEDIT_DADSB ] ) ;
end.
