{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 01/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGRECUPBUDGETPREV ()
Mots clefs ... : TOF;PGRECUPBUDGETPREV
*****************************************************************}
Unit UTofPGRecupBudgetPrevisionnel;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     sysutils,HCtrls,HEnt1,HMsgBox,UTOF,HTB97,UTob ;

Type
  TOF_PGRECUPBUDGETPREV = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    Trace : TListBox;
    Function RecupReporte(MillesimeS,MillesimeD:String):Boolean;
    Function RecupNonRealise(MillesimeS,MillesimeD:String):Boolean;
    Function RecupRefuse(MillesimeS,MillesimeD:String):Boolean;
    Function RecupForfait(MillesimeS,MillesimeD:String):Boolean;
    Function RecupPlafond(MillesimeS,MillesimeD:String):Boolean;
    Function RecupSalaires(MillesimeS,MillesimeD:String):Boolean;
    procedure TraiteRecup(Sender:TObject);
    Function CalculRang(MillesimeR,StageR,EtablissementR:String):Integer;
  end ;

Implementation

procedure TOF_PGRECUPBUDGETPREV.OnArgument (S : String ) ;
var Millesime,MillesimePrec:String;
    BTraitement:TToolBarButton97;
begin
  Inherited ;
        Millesime:=readTokenPipe(S,'');
        SetControlProperty('MILLESIMEDESTINATION','Value',Millesime);
        If IsNumeric(Millesime) Then
        begin
                MillesimePrec:=IntToStr(StrToInt(Millesime)-1);
                SetControlProperty('MILLESIMESOURCE','Value',MillesimePrec);
        end;
        BTraitement:=TToolBarButton97(GetControl('BTRAITEMENT'));
        If BTraitement<>Nil Then BTraitement.OnClick:=TraiteRecup;
end ;

Function TOF_PGRECUPBUDGETPREV.RecupReporte(MillesimeS,MillesimeD:String):Boolean;
var Tobreporte,T:Tob;
    Q:TQuery;
    Stage,Etablissement:String;
    Rang:Integer;
begin
        Trace.Items.Add('Récupération des inscriptions reportées :');
        Q:=OpenSQL('SELECT * FROM INSCFORMATION WHERE PFI_MILLESIME="'+MillesimeS+'"'+
          ' AND PFI_ETATINSCFOR="REP"',True);
        Tobreporte:=Tob.Create('INSCFIORMATION',Nil,-1);
        Tobreporte.LoaddetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        T:=Tobreporte.FindFirst(['PFI_MILLESIME'],[MillesimeS],False);
        While T<>Nil do
        begin
                Stage:=T.GetValue('PFI_CODESTAGE');
                Etablissement:=T.GetValue('PFI_ETABLISSEMENT');
                Rang:=CalculRang(MillesimeD,Stage,Etablissement);
                T.ChangeParent(Tobreporte,-1);
                T.PutValue('PFI_MILLESIME',MillesimeD);
                T.PutValue('PFI_RANG',Rang);
                T.PutValue('PFI_MILLESIME',MillesimeD);
                T.PutValue('PFI_DATEACCEPT',IDate1900);
                T.PutValue('PFI_REFUSELE',IDate1900);
                T.PutValue('PFI_ETATINSCFOR','ATT');
                T.PutValue('PFI_MOTIFETATINSC','');
                T.PutValue('PFI_REPORTELE',IDate1900);
                T.InsertDB(Nil,False);
                Trace.Items.Add('Libellé : '+T.GetValue('PFI_LIBELLE')+', Etablissement : '+T.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+T.GetValue('PFI_LIBEMPLOIFOR'));
                T:=Tobreporte.FindNext(['PFI_MILLESIME'],[MillesimeS],False);
        end;
        Tobreporte.Free;
        Trace.Items.Add('Fin de la récupération des inscriptions reportées');
        Result:=True;
end;

Function TOF_PGRECUPBUDGETPREV.RecupNonRealise(MillesimeS,MillesimeD:String):Boolean;
var TobInsc,TobFormationsBN,TobFormations,TI,TB,TF:Tob;
    QE,QI,QF:TQuery;
    DateDebut,DateFin:TDateTime;
    Salarie,Etab,LibEmploi,Stage:String;
    NbSal,NbInsc,Rang:Integer;
begin
        Trace.Items.Add('Récupération des inscriptions non réalisées :');
        QE:=OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeS+'"',True);
        Datedebut:=Date;
        DateFin:=Date;
        If not QE.eof then
        begin
                DateDebut:=QE.FindField('PFE_DATEDEBUT').AsDateTime;
                DateFin:=QE.FindField('PFE_DATEFIN').AsDateTime;
        end;
        Ferme(QE);

        QI:=OpenSQL('SELECT * FROM INSCFORMATION WHERE PFI_MILLESIME="'+MillesimeS+'"'+
        ' AND PFI_DATEACCEPT<>"'+UsDateTime(IDate1900)+'" AND PFI_REPORTELE="'+UsDateTime(IDate1900)+'"'+
        ' AND PFI_REFUSELE="'+UsDateTime(IDate1900)+'"',True);
        TobInsc:=Tob.Create('INSCFORMATION',Nil,-1);
        TobInsc.LoadDetailDB('INSCFORMATION','','',QI,False);
        Ferme(QI);

        QF:=OpenSQL('SELECT PFO_LIBEMPLOIFOR,PFO_ETABLISSEMENT,PFO_SALARIE FROM FORMATIONS '+
        'WHERE PFO_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'"'+
        ' AND PFO_EFFECTUE="X" AND EXISTS (SELECT PFI_SALARIE FROM INSCFORMATION WHERE PFI_MILLESIME="'+MillesimeS+'"'+
        ' AND PFI_SALARIE=PFO_SALARIE)',True);
        TobFormationsBN:=Tob.Create('Les formations BN',Nil,-1);
        TobFormationsBN.LoadDetailDB('FORMATIONS','','',QF,False);
        Ferme(QF);

        QF:=OpenSQL('SELECT PFO_LIBEMPLOIFOR,PFO_ETABLISSEMENT,COUNT (PFO_SALARIE) AS NBSAL FROM FORMATIONS '+
        'WHERE PFO_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'"'+
        ' AND PFO_EFFECTUE="X" AND NOT EXISTS (SELECT PFI_SALARIE FROM INSCFORMATION WHERE PFI_MILLESIME="'+MillesimeS+'"'+
        ' AND PFI_SALARIE=PFO_SALARIE) GROUP BY PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR',True);
        TobFormations:=Tob.Create('Les formations',Nil,-1);
        TobFormations.LoadDetailDB('FORMATIONS','','',QF,False);
        Ferme(QF);

        TI:=TobInsc.FindFirst([''],[''],False);
        While TI<>Nil do
        begin
                Salarie:=TI.GetValue('PFI_SALARIE');
                Etab:=TI.GetValue('PFI_ETABLISSEMENT');
                LibEmploi:=TI.GetValue('PFI_LIBEMPLOIFOR');
                If Salarie<>'' Then
                begin
                        TB:=TobFormationsBN.FindFirst(['PFO_SALARIE'],[Salarie],False);
                        If TB=Nil Then
                        begin
                                TI.ChangeParent(TobInsc,-1);
                                Stage:=TI.GetValue('PFI_CODESTAGE');
                                Rang:=CalculRang(MillesimeD,Stage,Etab);
                                TI.PutValue('PFI_RANG',Rang);
                                TI.PutValue('PFI_MILLESIME',MillesimeD);
                                TI.PutValue('PFI_DATEACCEPT',IDate1900);
                                TI.PutValue('PFI_REFUSELE',IDate1900);
                                TI.PutValue('PFI_REPORTELE',IDate1900);
                                TI.PutValue('PFI_REPORTEPAR','');
                                TI.PutValue('PFI_REFUSEPAR','');
                                TI.PutValue('PFI_REPORTFORM','');
                                TI.PutValue('PFI_REFUSFORM','');
                                TI.InsertDB(Nil,False);
                                Trace.Items.Add('Libellé : '+TI.GetValue('PFI_LIBELLE')+', Etablissement : '+TI.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+TI.GetValue('PFI_LIBEMPLOIFOR'));
                        end;
                end
                Else
                begin
                        TF:=TobFormations.FindFirst(['PFO_ETABLISSEMENT','PFO_LIBEMPLOIFOR'],[Etab,LibEmploi],False);
                        If TF<>Nil Then
                        begin
                                NbSal:=TF.GetValue('NBSAL');
                                While (TF<>Nil) and (NbSal<=0) do
                                begin
                                        TF:=TobFormations.FindFirst(['PFO_ETABLISSEMENT','PFO_LIBEMPLOIFOR'],[Etab,LibEmploi],False);
                                        NbSal:=TF.GetValue('NBSAL');
                                end;
                                NbInsc:=TI.GetValue('PFI_NBINSC');
                                If NbInsc>NbSal Then
                                begin
                                        Stage:=TI.GetValue('PFI_CODESTAGE');
                                        TI.ChangeParent(TobInsc,-1);
                                        Rang:=CalculRang(MillesimeD,Stage,Etab);
                                        TI.PutValue('PFI_RANG',Rang);
                                        TI.PutValue('PFI_MILLESIME',MillesimeD);
                                        TI.PutValue('PFI_NBISNC',NbInsc-NbSal);
                                        TI.PutValue('PFI_DATEACCEPT',IDate1900);
                                        TI.PutValue('PFI_REFUSELE',IDate1900);
                                        TI.PutValue('PFI_REPORTELE',IDate1900);
                                        TI.PutValue('PFI_REPORTEPAR','');
                                        TI.PutValue('PFI_REFUSEPAR','');
                                        TI.PutValue('PFI_REPORTFORM','');
                                        TI.PutValue('PFI_REFUSFORM','');
                                        TI.InsertDB(Nil,False);
                                        Trace.Items.Add('Libellé : '+TI.GetValue('PFI_LIBELLE')+', Etablissement : '+TI.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+TI.GetValue('PFI_LIBEMPLOIFOR'));
                                end
                                Else
                                begin
                                        TF.ChangeParent(TobFormations,-1);
                                        TF.PutValue('NBSAL',NbSal-NbInsc);
                                end;
                        end
                        Else
                        begin
                                Stage:=TI.GetValue('PFI_CODESTAGE');
                                TI.ChangeParent(TobInsc,-1);
                                Rang:=CalculRang(MillesimeD,Stage,Etab);
                                TI.PutValue('PFI_MILLESIME',MillesimeD);
                                TI.PutValue('PFI_RANG',Rang);
                                TI.PutValue('PFI_DATEACCEPT',IDate1900);
                                TI.PutValue('PFI_REFUSELE',IDate1900);
                                TI.PutValue('PFI_REPORTELE',IDate1900);
                                TI.PutValue('PFI_REPORTEPAR','');
                                TI.PutValue('PFI_REFUSEPAR','');
                                TI.PutValue('PFI_REPORTFORM','');
                                TI.PutValue('PFI_REFUSFORM','');
                                TI.InsertDB(Nil,False);
                                Trace.Items.Add('Libellé : '+TI.GetValue('PFI_LIBELLE')+', Etablissement : '+TI.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+TI.GetValue('PFI_LIBEMPLOIFOR'));
                        end;
                end;
                TI:=TobInsc.FindNext([''],[''],False);
        end;
        Trace.Items.Add('Fin de la récupération des inscriptions non réalisées.');
        TobFormations.Free;
        TobFormationsBN.Free;
        TobInsc.Free;
        Result:=True;
end;

Function TOF_PGRECUPBUDGETPREV.RecupRefuse(MillesimeS,MillesimeD:String):Boolean;
var TobRefuse,T:Tob;
    Q:TQuery;
    Stage,Etablissement:String;
    Rang:Integer;
begin
        Trace.Items.Add('Récupération des inscriptions refusées :');
        Q:=OpenSQL('SELECT * FROM INSCFORMATION WHERE PFI_MILLESIME="'+MillesimeS+'"'+
        ' AND PFI_ETATINSCFOR="REP"',True);
        TobRefuse:=Tob.Create('INSCFIORMATION',Nil,-1);
        TobRefuse.LoaddetailDB('INSCFORMATION','','',Q,False);
        Ferme(Q);
        T:=TobRefuse.FindFirst(['PFI_MILLESIME'],[MillesimeS],False);
        While T<>Nil do
        begin
                Stage:=T.GetValue('PFI_CODESTAGE');
                Etablissement:=T.GetValue('PFI_ETABLISSEMENT');
                Rang:=CalculRang(MillesimeD,Stage,Etablissement);
                T.ChangeParent(TobRefuse,-1);
                T.PutValue('PFI_MILLESIME',MillesimeD);
                T.PutValue('PFI_RANG',Rang);
                T.PutValue('PFI_ETATINSCFOR','ATT');
                T.PutValue('PFI_MOTIFETATINSC','');
                T.PutValue('PFI_REFUSELE',IDate1900);
                T.PutValue('PFI_REPORTELE',IDate1900);
                T.PutValue('PFI_DATEACCEPT',IDate1900);
                T.InsertDB(Nil,False);
                Trace.Items.Add('Libellé : '+T.GetValue('PFI_LIBELLE')+', Etablissement : '+T.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+T.GetValue('PFI_LIBEMPLOIFOR'));
                T:=TobRefuse.FindNext(['PFI_MILLESIME'],[MillesimeS],False);
        end;
        TobRefuse.Free;
        Trace.Items.Add('Fin de la récupération des inscriptions refusées');
        Result:=True;
end;

Function TOF_PGRECUPBUDGETPREV.RecupSalaires(MillesimeS,MillesimeD:String):Boolean;
var TobSalaires : Tob;
     Q : TQuery;
     i,Rep : Integer;
     LibEmploi : String;
begin
        If ExisteSQL('SELECT PSF_LIBEMPLOIFOR FROM SALAIREFORM WHERE PSF_MILLESIME="'+MillesimeD+'"') then
        begin
                Rep := PGIAsk('Attention vous avez saisie des salaires catégoriels pour le millésime '+MillesimeD+
                '#13#10voulez-vous poursuivre la récupération des salaries catégoriels',Ecran.Caption);
                If Rep <> MrYes then
                begin
                        Trace.Items.Add('Abandon de la récupération des salaries catégoriels');
                        Result := False;
                        Exit;
                end;
        end;
        Trace.Items.Add('Récupération des salaries catégoriels :');
        Q:=OpenSQL('SELECT * FROM SALAIREFORM WHERE PSF_MILLESIME="'+MillesimeS+'"',True);
        TobSalaires:=Tob.Create('SALAIREFORM',Nil,-1);
        TobSalaires.LoaddetailDB('SALAIREFORM','','',Q,False);
        Ferme(Q);
        For i := 0 to TobSalaires.Detail.Count - 1 do
        begin
                LibEmploi := TobSalaires.Detail[i].GetValue('PSF_LIBEMPLOIFOR');
                TobSalaires.Detail[i].PutValue('PSF_MILLESIME',MillesimeD);
                TobSalaires.InsertOrUpdateDB(False);
                Trace.Items.Add('Libellé emploi: '+RechDom('PGLIBEMPLOI',LibEmploi,False));
        end;
        TobSalaires.Free;
        Trace.Items.Add('Fin de la récupération des salaries catégoriels');
        Result:=True;
end;

procedure TOF_PGRECUPBUDGETPREV.TraiteRecup(Sender:TObject);
var Combo:THValComboBox;
    MillesimeSource,MillesimeDestination:STring;
begin
        Combo:=THValComboBox(GetControl('MILLESIMESOURCE'));
        MillesimeSource:=Combo.Value;
        Combo:=THValComboBox(GetControl('MILLESIMEDESTINATION'));
        MillesimeDestination:=Combo.Value;
        SetActiveTabSheet('PTRAITEMENT');
        Trace:= TListBox (GetControl ('LBTRAITES'));
        Trace.Items.Add('Début du traitement ');
        If GetControlText('CREPORTE')='X' Then RecupReporte(MillesimeSource,MillesimeDestination);
        If GetControlText('CREFUSE')='X' Then RecupRefuse(MillesimeSource,MillesimeDestination);
        If GetControlText('CNONEFFECTUE')='X' Then RecupNonRealise(MillesimeSource,MillesimeDestination);
        If GetControlText('CFORFAIT')='X' Then RecupForfait(MillesimeSource,MillesimeDestination);
        If GetControlText('CPLAFOND')='X' Then RecupPlafond(MillesimeSource,MillesimeDestination);
        If GetControlText('CSALAIREMOYEN')='X' Then RecupSalaires(MillesimeSource,MillesimeDestination);
        Trace.Items.Add('Fin du traitement ');
end;

Function TOF_PGRECUPBUDGETPREV.CalculRang(MillesimeR,StageR,EtablissementR:String):Integer;
var Q:TQuery;
    Rang:Integer;
begin
        Q:=OpenSQL('SELECT MAX(PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_MILLESIME="'+MillesimeR+'" '+
        'AND PFI_CODESTAGE="'+StageR+'" AND PFI_ETABLISSEMENT="'+EtablissementR+'"',True);
        if not Q.eof then Rang:=Q.FindField('RANG').AsInteger
        else Rang:=0;
        Rang:=Rang+1;
        Ferme(Q);
        Result:=Rang;
end;

Function TOF_PGRECUPBUDGETPREV.RecupForfait(MillesimeS,MillesimeD:String):Boolean;
var TobForfait,T : Tob;
    Q : TQuery;
begin
        Q := OpenSQL('SELECT * FROM FORFAITFORM WHERE PFF_MILLESIME="'+MillesimeS+'"',True);
        TobForfait := Tob.Create('FORFAITFORM',Nil,-1);
        TobForfait.LoaddetailDB('FORFAITFORM','','',Q,False);
        T := TobForfait.FindFirst(['PFF_MILLESIME'],[MillesimeS],False);
        Ferme(Q);
        While T<>Nil do
        begin
                T.ChangeParent(TobForfait,-1);
                T.PutValue('PFF_MILLESIME',MillesimeD);
                T.InsertOrUpdateDB(False);
                //Trace.Items.Add('Libellé : '+T.GetValue('PFI_LIBELLE')+', Etablissement : '+T.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+T.GetValue('PFI_LIBEMPLOIFOR'));
                T:=TobForfait.FindNext(['PFF_MILLESIME'],[MillesimeS],False);
        end;
        TobForfait.Free;
        Trace.Items.Add('Fin de la récupération des forfaits');
        Result:=True;
end;

Function TOF_PGRECUPBUDGETPREV.RecupPlafond(MillesimeS,MillesimeD:String):Boolean;
var TobPlafond,T : Tob;
    Q : TQuery;
begin
        Q := OpenSQL('SELECT * FROM FRAISSALPLAF WHERE PFP_MILLESIME="'+MillesimeS+'"',True);
        TobPlafond := Tob.Create('FRAISSALPLAF',Nil,-1);
        TobPlafond.LoaddetailDB('FRAISSALPLAF','','',Q,False);
        T := TobPlafond.FindFirst(['PFP_MILLESIME'],[MillesimeS],False);
        Ferme(Q);
        While T<>Nil do
        begin
                T.ChangeParent(TobPlafond,-1);
                T.PutValue('PFP_MILLESIME',MillesimeD);
                T.InsertOrUpdateDB(False);
                //Trace.Items.Add('Libellé : '+T.GetValue('PFI_LIBELLE')+', Etablissement : '+T.GetValue('PFI_ETABLISSEMENT')+', Libellé emploi :'+T.GetValue('PFI_LIBEMPLOIFOR'));
                T:=TobPlafond.FindNext(['PFP_MILLESIME'],[MillesimeS],False);
        end;
        TobPlafond.Free;
        Trace.Items.Add('Fin de la récupération des forfaits');
        Result:=True;
end;

Initialization
  registerclasses ( [ TOF_PGRECUPBUDGETPREV ] ) ;
end.
