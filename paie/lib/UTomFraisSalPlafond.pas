{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : FRAISSALPLAF (FRAISSALPLAF) et LIEUFORMATION
Mots clefs ... : TOM; FRAISSALPLAF ; LIEUFORMATION
*****************************************************************
PT1 24/11/2003 JL V_50 Modif pour CWAS
---- PH 10/08/2005 Suppression directive de compil $IFDEF AGL550B ----
---- JL 20/03/2006 modification clé annuaire ----
PT2 09/05/2007 FL V_720 FQ 11532 Ajout du code population pour la recherche des plafonds existants
PT3 05/05/2008 FL V_804 Dissimulation de l'établissement dans le cadre du partage formation
}
Unit UTomFraisSalPlafond;

Interface

Uses Controls,Classes,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,
{$ELSE}
     eFiche,eFichList,
{$ENDIF}
     sysutils,HCtrls,HEnt1,HMsgBox,UTOM,UTob,ed_tools,EntPaie,PGPopulOutils,PGOutilsFormation ; //PT2

Type
  TOM_FRAISSALPLAF = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
    Millesime:String;
    MajFrais:Boolean;
    InitMontant:Double;
    procedure MajFraisSaisies;
    end ;
  Type
  TOM_LIEUFORMATION = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnArgument (S : String)    ; override ;       //PT3
    end ;

Implementation
{TOM_FRAISSALPLAF}
procedure TOM_FRAISSALPLAF.OnNewRecord ;
begin
  Inherited ;
     SetField('PFP_MILLESIME',Millesime);
     SetField('PFP_PROVINCE','X');
     SetField('PFP_PARIS','X');
     SetField('PFP_ORGCOLLECT',-1);
end ;


procedure TOM_FRAISSALPLAF.OnLoadRecord ;
begin
  Inherited ;
     InitMontant:=GetField('PFP_MONTANT');
     TFFiche(Ecran).Caption:='Saisie des plafonds pour l''année '+Getfield('PFP_MILLESIME');
     UpdateCaption(TFFiche(Ecran));
end ;

procedure TOM_FRAISSALPLAF.OnUpdateRecord;
var Q:TQuery;
    TobPlafond,T:Tob;
    i,MaxOrdre:Integer;
    MessError,SQL:String;
begin
Inherited ;
     MajFrais:=False;
     MessError:='';

     { Vérification de la saisie }

     If GetField('PFP_MILLESIME')='' Then MessError:=MessError+'#13#10- Le millésime';
     If GetField('PFP_ORGCOLLECTGU')='' Then MessError:=MessError+'#13#10- L''organisme collecteur';
     If GetField('PFP_FRAISSALFOR')='' Then MessError:=MessError+'#13#10- Le type de frais';
     If (VH_Paie.PGForGestPlafByPop) Then If GetField('PFP_POPULATION')='' Then MessError:=MessError+'#13#10- La population'; //PT2
     If (GetField('PFP_PROVINCE')='-') and (Getfield('PFP_PARIS')='-') then MessError:=MessError+'#13#10- Plafond pour Province et/ou pour Paris';
     If MessError<>'' Then
     begin
        PGIBox('Vous devez renseigner la (ou les) information(s) suivante(s) :'+MessError,Ecran.Caption);
        LastError:=1;
        Exit;
     end;
     //PT2 - Début
     { Si la gestion des plafonds s'effectue par population et que celle propre à la formation
       n'est pas valide, il faut interdire la mise à jour. }
     If VH_Paie.PGForGestPlafByPop And Not CanUsePopulation(TYP_POPUL_FORM_PREV) Then
     Begin
           PGIBox('Attention, les populations de formation ne sont pas valides. #13#10'+
                  'La mise à jour des plafonds ne peut être effectuée.',Ecran.Caption);
           LastError := 1;
           Exit;
     End;
     //PT2 - Fin

     { Vérification de la cohérence : les frais saisis ne peuvent pas déjà exister }

     SQL := 'SELECT PFP_ORDRE,PFP_PROVINCE,PFP_PARIS,PFP_POPULATION FROM FRAISSALPLAF WHERE PFP_MILLESIME="'+GetField('PFP_MILLESIME')+'"'+
            ' AND PFP_ORGCOLLECTGU="'+GetField('PFP_ORGCOLLECTGU')+'" AND PFP_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'"' +
            ' AND PFP_ORDRE<>"'+IntToStr(GetField('PFP_ORDRE'))+'"';

     Q:=OpenSQL(SQL,True);
     TobPlafond:=Tob.Create('Les palfonds',Nil,-1);
     TobPlafond.LoadDetailDB('FRAISSALFORM','','',Q,False);
     Ferme(Q);
     If GetField('PFP_PARIS')='X' Then
     begin
        //PT2 - Début
        If (VH_Paie.PGForGestPlafByPop) Then
           T:=TobPlafond.FindFirst(['PFP_PARIS','PFP_POPULATION'],['X',GetField('PFP_POPULATION')],False)
        Else
           T:=TobPlafond.FindFirst(['PFP_PARIS','PFP_POPULATION'],['X',''],False);
        //PT2 - Fin
        If T<>Nil Then
        begin
           PGIBox('Il existe déjà un plafond pour Paris pour ce type de frais. Création impossible',TFFiche(Ecran).Caption);
           SetField('PFP_PARIS','-');
           TobPlafond.Free;
           LastError:=1;
           Exit;
        end;
     end;
     If GetField('PFP_PROVINCE')='X' Then
     begin
        //PT2 - Début
        If (VH_Paie.PGForGestPlafByPop) Then
           T:=TobPlafond.FindFirst(['PFP_PROVINCE','PFP_POPULATION'],['X',GetField('PFP_POPULATION')],False)
        Else
           T:=TobPlafond.FindFirst(['PFP_PROVINCE','PFP_POPULATION'],['X',''],False);
        //PT2 - Fin
        If T<>Nil Then
        begin
           PGIBox('Il existe déjà un plafond pour la province pour ce type de frais. Création impossible',TFFiche(Ecran).Caption);
           SetField('PFP_PROVINCE','-');
           TobPlafond.Free;
           LastError:=1;
           Exit;
       end;
     end;
     MaxOrdre:=0;
     For i:=0 to TobPlafond.Detail.Count - 1 do
     begin
          If TobPlafond.Detail[i].GetValue('PFP_ORDRE')>MaxOrdre Then MaxOrdre:=TobPlafond.Detail[i].GetValue('PFP_ORDRE');
     end;
     MaxOrdre:=MaxOrdre+1;
     SetField('PFP_ORDRE',MaxOrdre);
     TobPlafond.Free;

     { Recherche, dans le cas où le montant des frais a été modifié, s'il existe des frais saisis pour les salariés }

     If (InitMontant<>0) And (InitMontant<>GetField('PFP_MONTANT')) then
     begin
          If (GetField('PFP_PARIS')='X') and (GetField('PFP_PROVINCE')='-') then
          SQL:='SELECT PFS_MONTANT FROM FRAISSALFORM LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_MILLESIME=PSS_MILLESIME AND PFS_ORDRE=PSS_ORDRE'+
               ' LEFT JOIN LIEUFORMATION ON PLF_LIEUFORM=PSS_LIEUFORM'+
               ' WHERE PFS_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'" AND '+
               'PFS_MILLESIME="'+GetField('PFP_MILLESIME')+'" AND PLF_PLAFPARIS="X" AND PSS_ORGCOLLECTSGU="'+GetField('PFP_ORGCOLLECTGU')+'"';
          If (GetField('PFP_PARIS')='-') and (GetField('PFP_PROVINCE')='X') then
          SQL:='SELECT PFS_MONTANT FROM FRAISSALFORM LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_MILLESIME=PSS_MILLESIME AND PFS_ORDRE=PSS_ORDRE'+
               ' LEFT JOIN LIEUFORMATION ON PLF_LIEUFORM=PSS_LIEUFORM'+
               ' WHERE PFS_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'" AND '+
               'PFS_MILLESIME="'+GetField('PFP_MILLESIME')+'" AND PLF_PLAFPARIS="-" AND PSS_ORGCOLLECTSGU="'+GetField('PFP_ORGCOLLECTGU')+'"';
          If (GetField('PFP_PARIS')='X') and (GetField('PFP_PROVINCE')='X') then
          SQL:='SELECT PFS_MONTANT FROM FRAISSALFORM LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_MILLESIME=PSS_MILLESIME AND PFS_ORDRE=PSS_ORDRE'+
               ' WHERE PFS_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'" AND '+
               'PFS_MILLESIME="'+GetField('PFP_MILLESIME')+'" AND PSS_ORGCOLLECTSGU="'+GetField('PFP_ORGCOLLECTGU')+'"';

          //PT2 - Début
          If (VH_Paie.PGForGestPlafByPop) Then
               SQL := SQL + ' AND PFS_POPULATION="'+GetField('PFP_POPULATION')+'"';
          //PT2 - Fin
          
          If ExisteSQL(SQL) then
          begin 
               Case PGIAskCancel('Attention, il existe des frais saisis pour l''année '+GetField('PFP_MILLESIME')+ ' avec l''ancien plafond de '+FloatToStr(InitMontant)+
                                 '.#13#10 Voulez vous mettre à jour ces frais avec le nouveau plafond de '+FloatToStr(getField('PFP_MONTANT'))+' ?',Ecran.Caption) of
               mrYes : MajFrais:=True;
               mrCancel :
                    begin
                      LastError:=1;
                      exit;
                    end;
               end;
          end;
     end;
end;

procedure TOM_FRAISSALPLAF.OnAfterUpdateRecord;
begin
  Inherited ;
     If MajFrais=True then
     begin
          MajFraisSaisies;
     end;
end;

procedure TOM_FRAISSALPLAF.OnArgument ( S: String ) ;
begin
  Inherited ;
     Millesime:=ReadTokenPipe(S,';');
     TFFIche(Ecran).UniqueName:='PFP_MILLESIME,PFP_ORGCOLLECTGU,PFP_FRAISSALFOR,PFP_ORDRE';
     TFFiche(Ecran).DisabledMajCaption := True;

     //PT2 - Début
     // En cas de gestion par population, on affiche la combo et le libellé associé.
     If (VH_Paie.PGForGestPlafByPop) Then
     Begin
          SetControlVisible('TPFP_POPULATION', True);
          SetControlVisible('PFP_POPULATION',  True);
          // Mise à jour de la liste de la combo avec les populations actives
          SetControlProperty('PFP_POPULATION', 'Plus', ' AND PPC_PREDEFINI="'+GetPredefiniPopulation(TYP_POPUL_FORM_PREV)+'"');
     End;
     //PT2 - Fin
end ;

procedure TOM_FRAISSALPLAF.MajFraisSaisies;
var TobFrais:Tob;
    Q:TQuery;
    SQL:String;
    Qte,i:Integer;
    Montant,Plafond,MontantPlaf:Double;
begin
     Plafond:=GetField('PFP_MONTANT');

     { Recherche des frais saisis pour les salariés et mise à jour de ceux-ci s'ils dépassent le nouveau plafond }

     If (GetField('PFP_PARIS')='X') and (GetField('PFP_PROVINCE')='-') then
     SQL:='SELECT * FROM FRAISSALFORM LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_MILLESIME=PSS_MILLESIME AND PFS_ORDRE=PSS_ORDRE'+
          ' LEFT JOIN LIEUFORMATION ON PLF_LIEUFORM=PSS_LIEUFORM'+
          ' WHERE PFS_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'" AND '+
          'PFS_MILLESIME="'+GetField('PFP_MILLESIME')+'" AND PLF_PLAFPARIS="X" AND PSS_ORGCOLLECTSGU="'+GetField('PFP_ORGCOLLECTGU')+'"';
     If (GetField('PFP_PARIS')='-') and (GetField('PFP_PROVINCE')='X') then
     SQL:='SELECT * FROM FRAISSALFORM LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_MILLESIME=PSS_MILLESIME AND PFS_ORDRE=PSS_ORDRE'+
          ' LEFT JOIN LIEUFORMATION ON PLF_LIEUFORM=PSS_LIEUFORM'+
          ' WHERE PFS_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'" AND '+
          'PFS_MILLESIME="'+GetField('PFP_MILLESIME')+'" AND PLF_PLAFPARIS="-" AND PSS_ORGCOLLECTSGU="'+GetField('PFP_ORGCOLLECTGU')+'"';
     If (GetField('PFP_PARIS')='X') and (GetField('PFP_PROVINCE')='X') then
     SQL:='SELECT * FROM FRAISSALFORM LEFT JOIN SESSIONSTAGE ON PFS_CODESTAGE=PSS_CODESTAGE AND PFS_MILLESIME=PSS_MILLESIME AND PFS_ORDRE=PSS_ORDRE'+
          ' WHERE PFS_FRAISSALFOR="'+GetField('PFP_FRAISSALFOR')+'" AND '+
          'PFS_MILLESIME="'+GetField('PFP_MILLESIME')+'" AND PSS_ORGCOLLECTSGU="'+GetField('PFP_ORGCOLLECTGU')+'"';
     //PT2 - Début
     If (VH_Paie.PGForGestPlafByPop) Then SQL := SQL + ' AND PFS_POPULATION="'+GetField('PFP_POPULATION')+'"';
     //PT2 - Fin
     Q:=OpenSQL(SQL,True);
     TobFrais:=Tob.Create('FRAISALFORM',Nil,-1);
     TobFrais.LoadDetailDB('FRAISSALFORM','','',Q,False);
     Ferme(Q);

     InitMoveProgressForm (NIL,'Mise à jour des frais', 'Veuillez patienter SVP ...',TobFrais.Detail.Count-1,FALSE,TRUE);
     For i:=0 to TobFrais.Detail.Count-1 do
     begin
          if Plafond=0 then MontantPlaf:=TobFrais.Detail[i].GetValue('PFS_MONTANT')
          Else
          begin
               Qte:=TobFrais.Detail[i].GetValue('PFS_QUANTITE');
               Montant:=TobFrais.Detail[i].GetValue('PFS_MONTANT');
               MontantPlaf:=Arrondi(Qte*Plafond,2);
               If Montant<MontantPlaf then MontantPlaf:=Montant;
          end;
          TobFrais.Detail[i].PutValue('PFS_MONTANTPLAF',MontantPlaf);
          TobFrais.Detail[i].UpdateDB(False);
     end;
     TobFrais.Free;
     FiniMoveProgressForm;
end;

{TOM_LIEUFORMATION}
procedure TOM_LIEUFORMATION.OnNewRecord ;
begin
  Inherited ;
     SetField('PLF_PLAFPARIS','-');
end ;

procedure TOM_LIEUFORMATION.OnAfterUpdateRecord ;
begin
  Inherited ;
     AvertirTable('PGLIEUFORMATION');
end;

procedure TOM_LIEUFORMATION.OnDeleteRecord ;
begin
  Inherited ;
     If ExisteSQL('SELECT PST_LIEUFORM FROM STAGE WHERE PST_LIEUFORM="'+GetField('PLF_LIEUFORM')+'"') then
     begin
        PGIBox('Suppression impossible car il existe des formations pour ce lieu',Ecran.Caption);
        LastError:=1;
     end;
     If ExisteSQL('SELECT PFF_LIEUFORM FROM FORFAITFORM WHERE PFF_LIEUFORM="'+GetField('PLF_LIEUFORM')+'"') then
     begin
        PGIBox('Suppression impossible car il existe des forfaits pour ce lieu',Ecran.Caption);
        LastError:=1;
     end;
end;

//PT3
procedure TOM_LIEUFORMATION.OnArgument (S : String);
begin
  inherited;
    If PGBundleCatalogue Then
    Begin
        SetControlVisible ('PLF_ETABLISSEMENT', False);
        SetControlVisible ('TPLF_ETABLISSEMENT', False);
    End;
end;

Initialization
  registerclasses ( [ TOM_FRAISSALPLAF,TOM_LIEUFORMATION ] ) ;
end.
