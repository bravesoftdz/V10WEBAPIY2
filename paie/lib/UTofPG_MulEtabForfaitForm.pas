{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULETABFORFAIT ()
Mots clefs ... : TOF;PGMULETABFORFAIT
*****************************************************************
PT1  15/05/2007 V_720 FL FQ 13567 Gestion des duplications avec les populations
}

Unit UTofPG_MulEtabForfaitForm;

Interface

Uses Controls,Classes,
{$IFNDEF EAGLCLIENT}
      HDB,Mul,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     emul,eFiche,UtileAGL,MaineAgl,
{$ENDIF}
     sysutils,HCtrls,UTOF,UTob,HTB97,EntPaie,PGOutilsFormation,PGPopulOutils ; //PT1

Type
  TOF_PGMULETABFORFAIT = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Lieu,Millesime:String;
    procedure DuplicationForfait(Sender:TObject);
    procedure ChangeFrais(Sender:TObject);
    procedure MAJFrais(TobForfait: TOB; LesQuels : String; FraisRepas,FraisHeberg,FraisTrans : Double);  //PT1
  end ;

Implementation

procedure TOF_PGMULETABFORFAIT.OnLoad ;
var St:String;
begin
  Inherited ;

     //PT1 - Début
     If VH_Paie.PGForGestFraisByPop Then
          St := 'PPO_TYPEPOP LIKE "%'+TYP_POPUL_FORM_PREV+'%" AND PPC_PREDEFINI="'+GetPredefiniPopulation(TYP_POPUL_FORM_PREV)+'" AND PPC_POPULATION NOT IN (SELECT DISTINCT PFF_POPULATION '
     Else
          St := 'ET_ETABLISSEMENT NOT IN (SELECT PFF_ETABLISSEMENT ';
     //PT1 - Fin

     St := St + 'FROM FORFAITFORM WHERE PFF_LIEUFORM="'+Lieu+'" AND PFF_MILLESIME="'+Millesime+'"';

     If GetControlText('FRAISDUPLIQUE')='R' Then St := St + ' AND PFF_FRAISREPAS<>0';
     If GetControlText('FRAISDUPLIQUE')='H' Then St := St + ' AND PFF_FRAISHEBERG<>0';
     If GetControlText('FRAISDUPLIQUE')='D' Then St := St + ' AND PFF_FRAISTRANSP<>0';

     St := St + ')';

     SetControlText('XX_WHERE',St);
end ;

procedure TOF_PGMULETABFORFAIT.OnArgument (S : String ) ;
var FraisRepas,FraisHeberg,FraisTrans: String;
    BLance:TToolBarButton97;
    Combo:THValComboBox;
begin
  Inherited ;
     Lieu:=ReadTokenPipe(S,';');
     Millesime:=ReadTokenPipe(S,';');
     FraisRepas:=ReadTokenPipe(S,';');
     FraisHeberg:=ReadTokenPipe(S,';');
     FraisTrans:=ReadTokenPipe(S,';');

     SetControlText('FORFAITREPAS',FraisRepas);
     SetControlText('FORFAITHEBERG',FraisHeberg);
     SetControlText('FORFAITDEPLACE',FraisTrans);
     SetControlProperty('FRAISDUPLIQUE','Value','T');

     BLance:=TToolBarButton97(Getcontrol('BOuvrir'));
     If BLance<>Nil Then BLance.OnClick:=DuplicationForfait;

     Combo:=THValComboBox(GetControl('FRAISDUPLIQUE'));
     If Combo<>Nil Then Combo.OnChange:=ChangeFrais;

     Ecran.Caption:='Duplication des frais pour le lieu de formation : '+RechDom('PGLIEUFORMATION',Lieu,False);

     // Dans le cas des populations, on change la DBListe
     //PT1 - Début
     If VH_Paie.PGForGestFraisByPop Then
     Begin
          TFMul(Ecran).SetDBListe('PGPOPULATION');
     End;
     //PT1 - Fin
end ;

procedure TOF_PGMULETABFORFAIT.DuplicationForfait(Sender:TObject);
var TobForfait:Tob;
    Q:TQuery;
    LesQuels:String;
    i:Integer;
    {$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
    {$ELSE}
    Liste:THGrid;
    {$ENDIF}
    FraisRepas,FraisHeberg,FraisTrans:Double;
begin
     LesQuels    := GetControlText('FRAISDUPLIQUE');
     FraisRepas  := StrToFloat(GetControlText('FORFAITREPAS'));
     FraisHeberg := StrToFloat(GetControlText('FORFAITHEBERG'));
     FraisTrans  := StrToFloat(GetControlText('FORFAITDEPLACE'));

     {$IFNDEF EAGLCLIENT}
     Liste:=THDBGrid(GetControl('FLISTE'));
     {$ELSE}
     Liste:=THGrid(GetControl('FLISTE'));
     {$ENDIF}

     Q:=OpenSQL('SELECT * FROM FORFAITFORM '+
                ' WHERE PFF_LIEUFORM="'+Lieu+'" AND PFF_MILLESIME="'+Millesime+'"',True);
     TobForfait:=Tob.Create('FORFAITFORM',Nil,-1);
     TobForfait.LoaddetailDB('FORFAITFORM','','',Q,False);
     Ferme(Q);

     {$IFDEF EAGLCLIENT}
     if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
     {$ENDIF}

     // Cas 1 : Tous les éléments ont été sélectionnés
     If (Liste.AllSelected=TRUE) then
     begin
          // InitMove(TFmul(Ecran).Q.RecordCount,'');
          TFmul(Ecran).Q.First;
          while Not TFmul(Ecran).Q.EOF do
          begin
               MAJFrais(TobForfait,LesQuels,FraisRepas,FraisHeberg,FraisTrans);    //PT1
               TFmul(Ecran).Q.Next;
          end;
     end
     Else
     Begin
          // Cas 2 : Seuls certains éléments ont été sélectionnés
          for i:=0 to Liste.NbSelected-1 do
          begin
               Liste.GotoLeBOOKMARK(i);
               {$IFDEF EAGLCLIENT}
               TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
               {$ENDIF}
               MAJFrais (TobForfait,LesQuels,FraisRepas,FraisHeberg,FraisTrans); //PT1
          end;
     End;
     Liste.AllSelected:=False;
     TobForfait.Free;
     TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULETABFORFAIT.ChangeFrais(Sender:TObject);
begin
TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 15/05/2007 / PT1
Modifié le ... :   /  /    
Description .. : Crée les frais pour un établissement ou une population 
Suite ........ : donnée en fonction des frais existants
Mots clefs ... : POPULATION;ETABLISSEMENT
*****************************************************************}
procedure TOF_PGMULETABFORFAIT.MAJFrais(TobForfait: TOB; LesQuels : String; FraisRepas,FraisHeberg,FraisTrans : Double);
Var Pop,Etab : String;
    T: TOB;
begin
     If (VH_Paie.PGForGestFraisByPop) Then
     Begin
          Pop:=TFmul(Ecran).Q.FindField('PPC_POPULATION').AsString;
          T:=TobForfait.FindFirst(['PFF_POPULATION'],[Pop],False);
     End
     Else
     Begin
          Etab:=TFmul(Ecran).Q.FindField('ET_ETABLISSEMENT').AsString;
          T:=TobForfait.FindFirst(['PFF_ETABLISSEMENT'],[Etab],False);
     End;
     If T <> Nil Then
     begin
          // Dans le cas où la population ou l'établissement existe déjà, on ne modifie que les frais
          If (LesQuels='T') or (LesQuels='R') Then T.PutValue('PFF_FRAISREPAS',FraisRepas);
          If (LesQuels='T') or (LesQuels='H') Then T.PutValue('PFF_FRAISHEBERG',FraisHeberg);
          If (LesQuels='T') or (LesQuels='R') Then T.PutValue('PFF_FRAISTRANSP',FraisTrans);
          T.UpdateDB(False);
     end
     Else
     begin
          // Dans le cas où la population ou l'établissement n'est pas trouvé, on crée la TOB
          T:=Tob.Create('FORFAITFORM',TobForfait,-1);
          //PT1 - Début
          If (VH_Paie.PGForGestFraisByPop) Then
          Begin
               T.PutValue('PFF_POPULATION',Pop);
               T.PutValue('PFF_ETABLISSEMENT','---');
          End
          Else
          Begin
               T.PutValue('PFF_ETABLISSEMENT',Etab);
               T.PutValue('PFF_POPULATION','---');
               T.PutValue('PFF_CODEPOP','---');
          End;
          //PT1 - Fin
          T.PutValue('PFF_MILLESIME',Millesime);
          T.PutValue('PFF_LIEUFORM',Lieu);
          If (LesQuels='T') or (LesQuels='R') Then T.PutValue('PFF_FRAISREPAS',FraisRepas);
          If (LesQuels='T') or (LesQuels='H') Then T.PutValue('PFF_FRAISHEBERG',FraisHeberg);
          If (LesQuels='T') or (LesQuels='R') Then T.PutValue('PFF_FRAISTRANSP',FraisTrans);
          T.InsertOrUpdateDB(False);
     end;
end;

Initialization
  registerclasses ( [ TOF_PGMULETABFORFAIT ] ) ;
end.
