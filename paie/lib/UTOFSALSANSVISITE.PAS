{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 17/07/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF Edition des salariés sans visite médicale
Mots clefs ... : PAIE;MEDECINETRAVAIL
*****************************************************************
PT- 1   V563 04/12/2001 JL : Ajout de la gestion des dates Fiche bug n°375
PT- 2   V585 08/10/2002 JL : Modification de la gestion période
PT- 3   V_50 07/11/2003 JL : FQ 10971 Ajout critères pour exclurre salariés sorties
PT- 4   V_65 04/03/2006 JL : FQ 12893 Ajout champs libres pour gestion surveillance médicale renforcée
PT- 5 : 31/01/2007 V_80 FC : Alimentation du champ XX_WHERE au lieu de passer la requête entière
                             Renommage des champs critères Etablissement XX_VARETABx en PSA_ETABLISSEMENT
                             pour permettre la récupération automatique des habilitations (standardisation)
PT- 6 : V-80 10/09/2007 NA : FQ 14733 Prise en compte des sélections et du tri en utilisant XX_WHERE et XX_ORDERBY
 }
Unit UTOFSALSANSVISITE ;

Interface

Uses StdCtrls, Controls, Classes, sysutils,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,
{$ELSE}
     eQRS1,UtileAGL,UTOB,
{$ENDIF}
     HCtrls, HEnt1,  UTOF,ParamDat,Spin,EntPaie,HTB97,P5Def ;

Type
  TOF_SALSANSVISITE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure DateElipsisclick(Sender: TObject); // PT- 1
    procedure periode(Sender:TObject); // PT- 1
    procedure AccesDateSortie(Sender : TObject);
    procedure RequeteEtat(Sender : TObject); // PT- 6
  end ;

Implementation
// Confection de la requete dans la fiche QRS1 en fonction des critères posés sur la fiche
// deb  PT_ 6
//procedure TOF_SALSANSVISITE.OnUpdate ;
  procedure TOF_SALSANSVISITE.RequeteEtat(Sender : TObject) ;
var SQL,OrderBy,StWhere : String;
    DateDebut,DateFin : TDateTime;
    TNbMois : TSpinEdit;
    NbMois : Integer;
begin
  Inherited ;
        TNbMois := TSpinEdit(GetControl('NBMOIS'));
        NbMois := TNbMois.Value;
        NbMois := - NbMois;
        DateFin := StrToDate(GetcontrolText('DATEFIN'));
        DateDebut := PlusMois(Datefin,NbMois);
        SetControlText('DATEDEBUT',DateToStr(Datedebut));


        // DEBUT PT- 1
        If GetCheckBoxState('CDATE') = CbChecked then

                // DEBUT PT- 2

                { PT5  SQL := 'SELECT PSA_PRENOM,PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_DATEENTREE FROM SALARIES'+
                       ' WHERE PSA_ETABLISSEMENT>="'+GetControlText('XX_VARETAB1')+'" AND PSA_ETABLISSEMENT<="'+GetControlText('XX_VARETAB2')+'"'+
                       ' AND PSA_DATEENTREE<"'+UsDateTime(DateFin)+'"'+
                       ' AND PSA_SALARIE NOT IN (SELECT PVM_SALARIE FROM VISITEMEDTRAV WHERE '+
                       'PVM_DATEVISITE>="'+UsDateTime(DateDebut)+'" AND PVM_DATEVISITE<="'+UsDateTime(DateFin)+'")'+StWhere;}
                Stwhere := ' PSA_DATEENTREE<"'+UsDateTime(DateFin)+'"'+
                           ' AND PSA_SALARIE NOT IN (SELECT PVM_SALARIE FROM VISITEMEDTRAV WHERE '+
                           'PVM_DATEVISITE>="'+UsDateTime(DateDebut)+'" AND PVM_DATEVISITE<="'+UsDateTime(DateFin)+'")'
               //FIN PT- 2

        else
        { PT5       SQL := 'SELECT PSA_PRENOM,PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_DATEENTREE FROM SALARIES'+
                    ' WHERE PSA_ETABLISSEMENT>="'+GetControlText('XX_VARETAB1')+'" AND PSA_ETABLISSEMENT<="'+GetControlText('XX_VARETAB2')+'"'+
                    ' AND PSA_SALARIE NOT IN (SELECT DISTINCT(PVM_SALARIE) FROM VISITEMEDTRAV)'+StWhere;}
                Stwhere := ' PSA_SALARIE NOT IN (SELECT DISTINCT(PVM_SALARIE) FROM VISITEMEDTRAV)';
        // FIN PT- 1

        If GetCheckBoxState('CDATESORTIE') = CbChecked then  //PT- 3

                SQL := stwhere + ' AND (PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE >"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")'

        else    SQL := stwhere;

        If GetCheckBoxState('CETAB') = CbChecked Then SetControlText('XX_RUPTURE1','PSA_ETABLISSEMENT')
        Else SetControlText('XX_RUPTURE1','');
     // deb PT-6
     //   If GetCheckBoxState('CTRI')= CbChecked Then OrderBy:=' ORDER BY PSA_ETABLISSEMENT,PSA_LIBELLE'
     //   Else OrderBy:=' ORDER BY PSA_ETABLISSEMENT,PSA_SALARIE';
          If GetCheckBoxState('CTRI')= CbChecked Then
          begin
            If GetCheckBoxState('CETAB') = CbChecked then orderby := ' PSA_LIBELLE'
            else OrderBy:=' PSA_ETABLISSEMENT,PSA_LIBELLE';
          end
          Else
          begin
            If GetCheckBoxState('CETAB') = CbChecked then OrderBy:=' PSA_SALARIE'
            else OrderBy:=' PSA_ETABLISSEMENT,PSA_SALARIE';
          end;

     // fin PT-6
        SetControlText('XX_WHERE', SQL);        //PT5
        SetControlText('XX_ORDERBY', OrderBy);  //PT5
        // TFQRS1(Ecran).WhereSQL := SQL+OrderBy; //PT5
end ;

procedure TOF_SALSANSVISITE.OnUpdate ;
begin
Inherited ;
end;

// fin pt6

// Bornage des etablissements par defaut
procedure TOF_SALSANSVISITE.OnArgument (S : String ) ;
var     Q1,Q2 : TQuery;
        Etab1,Etab2 : String;
        Check : TCheckBox;
        Edit : THEdit;
        NBMS : TSpinEdit;
        Num : Integer;
begin
Inherited ;

        Q1:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS ORDER BY ET_ETABLISSEMENT',True);
        Etab1:='';
        If not Q1.eof then Etab1:=Q1.FindField('ET_ETABLISSEMENT').AsString;// // PortageCWAS
        Ferme(Q1);
        SetControlText('PSA_ETABLISSEMENT',Etab1);  //PT5
        Q2:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS ORDER BY ET_ETABLISSEMENT DESC',True);
        Etab2:='';
        if not Q2.eof then Etab2:=Q2.FindField('ET_ETABLISSEMENT').AsString; // // PortageCWAS
        Ferme(Q2);
        SetControlText('PSA_ETABLISSEMENT_',Etab2);  //PT5
        Edit := THEdit(GetControl('DATEDEBUT'));
        If Edit <> NIL Then                    // PT- 1
        begin
                Edit.Enabled := False;
                Edit.Text := DateToStr(idate1900);
                Edit.OnElipsisClick := DateElipsisclick;
        end;
        Edit := THEdit(GetControl('DATEFIN'));
        If Edit <>NIL Then
        begin
                SetControlEnabled('NBMOIS',False);
                Edit.Enabled := False;
                Edit.Text:=DateToStr(idate1900);
                Edit.OnElipsisClick := DateElipsisclick;
        end;
        Check := TCheckBox(GetControl('CDATE'));
        If Check <> NIL Then Check.OnClick := Periode;
        // DEBUT PT- 3
        Edit := THEdit(GetControl('DATESORTIE'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        SetControlChecked('CDATESORTIE',True);
        SetControlText('DATESORTIE',DateToStr(V_PGI.DateEntree));
        // Deb PT_ 6
        NbMS := TSpinEdit(GetControl('NBMOIS'));
        If NbMS <> Nil then NbMS.OnChange:= RequeteEtat;
        Check := TCheckBox(GetControl('CETAB'));
        If Check <> NIL Then Check.OnClick := RequeteEtat;
        Check := TCheckBox(GetControl('CTRI'));
        If Check <> NIL Then Check.OnClick := RequeteEtat;
        // fin PT- 6

        Check := TCheckBox(GetControl('CDATESORTIE'));
        If Check <> Nil then Check.OnClick := AccesDateSortie;
        // FIN PT- 3
        // DEBUT PT- 4
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
             if Num >4 then Break;
             VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        For Num := 1 to VH_Paie.PgNbCombo do
        begin
             if Num >4 then Break;
             VisibiliteChampLibreSal(IntToStr(Num),GetControl ('PSA_LIBREPCMB'+IntToStr(Num)),GetControl ('TPSA_LIBREPCMB'+IntToStr(Num)));
        end;
        // FIN PT- 4
        RequeteEtat(nil);   // PT_ 6

end ;


procedure TOF_SALSANSVISITE.periode(Sender:TObject); // PT- 1
begin
        If GetCheckBoxState('CDATE') = CbChecked Then
        begin
                SetControlEnabled('DATEDEBUT',True);
                SetControlEnabled('DATEFIN',True);
                SetControlEnabled('NBMOIS',True);
                SetControlText('DATEDEBUT',DateToStr(Date));
                SetControlText('DATEFIN',DateToStr(Date));
        end
        Else
        begin
                SetControlEnabled('DATEDEBUT',False);
                SetControlEnabled('DATEFIN',False);
                SetControlEnabled('NBMOIS',False);
                SetControlText('DATEDEBUT',DateToStr(idate1900));
                SetControlText('DATEFIN',DateToStr(IDate1900));
        end;
end;

procedure TOF_SALSANSVISITE.DateElipsisclick(Sender: TObject); // PT- 1
var key : char;
begin
        key := '*';
        ParamDate (Ecran, Sender, Key);
        requeteEtat(nil); // PT- 6
end;

procedure TOF_SALSANSVISITE.AccesDateSortie(Sender : TObject);
begin
        If GetCheckBoxState('CDATESORTIE') = CbChecked then SetControlEnabled('DATESORTIE',True)
        else SetControlEnabled('DATESORTIE',False);
        RequeteEtat(nil);  // PT_ 6
end;

Initialization
  registerclasses ( [ TOF_SALSANSVISITE ] ) ;
end.
