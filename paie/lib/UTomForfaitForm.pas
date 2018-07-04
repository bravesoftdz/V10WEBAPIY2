{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/08/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : FORFAITFORM (FORFAITFORM)
Mots clefs ... : TOM;FORFAITFORM
*****************************************************************
PT1 30/06/2006 V_70 JL FQ 13337 1 repas au lieu de 2 si pas d'hébergement
PT2 09/05/2007 V_720 FL FQ 13567 Gestion des populations
}
Unit UTomForfaitForm;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,Fe_Main,
{$ELSE}
     eFiche,eFichList,MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,HTB97,EntPaie,PGOutilsFormation,PGPopulOutils ; //PT2

Type
  TOM_FORFAITFORM = Class (TOM)
    procedure OnLoadRecord               ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    Private
    Millesime:String;
    procedure BDupliquerClick(Sender:TObject);
    end ;

Implementation

procedure TOM_FORFAITFORM.OnLoadRecord ;
begin
  Inherited ;
     TFFiche(Ecran).caption:='Saisie des forfaits pour le prévisionnel : '+GetField('PFF_MILLESIME');
     UpdateCaption(TFFiche(Ecran));
end ;

procedure TOM_FORFAITFORM.OnNewRecord ;
begin
  Inherited ;
     SetField('PFF_MILLESIME',Millesime);
     //PT2 - Début
     If (VH_Paie.PGForGestFraisByPop) Then
     Begin
          SetField('PFF_ETABLISSEMENT','---');
          SetField('PFF_CODEPOP','---');
     End
     Else
     Begin
          SetField('PFF_POPULATION','---');
          SetField('PFF_CODEPOP','---');
     End;
     //PT2 - Fin
end ;

procedure TOM_FORFAITFORM.OnUpdateRecord;
var St,MessError:String;
    rep    : Integer;
    TobInsc:Tob;
    FraisH,FraisR,FraisT,NbJour,TotalFrais:Double;
    NbInsc,i:Integer;
begin
     MessError:='';

     { Vérification de la saisie }

     //PT2 - Début
     If (VH_Paie.PGForGestFraisByPop) Then
     Begin
          If GetField('PFF_POPULATION')='' then MessError:=MessError+'#13#10- la population';
     End
     Else
     Begin
     //PT2 - Fin
          If GetField('PFF_ETABLISSEMENT')='' then MessError:=MessError+'#13#10- l''établissement';
     End;
     If GetField('PFF_LIEUFORM')='' then MessError:=MessError+'#13#10- le lieu de formation';
     If GetField('PFF_MILLESIME')='' then MessError:=MessError+'#13#10- le millésime';
     If MessError<>'' then
     begin
        PGIBox('Vous devez renseigner la (ou les) information(s) suivante(s) :'+MessError,Ecran.Caption);
        LastError:=1;
        Exit;
     end;

     { Mise à jour des frais existants }

     //PT2 - Début
     St:='SELECT * FROM INSCFORMATION WHERE PFI_MILLESIME="'+GetField('PFF_MILLESIME')+'" AND PFI_LIEUFORM="'+GetField('PFF_LIEUFORM')+'"';
     If (VH_Paie.PGForGestFraisByPop) Then
          St := St + ' AND (PFI_SALARIE IN (SELECT PNA_SALARIE FROM SALARIEPOPUL WHERE PNA_TYPEPOP="'+TYP_POPUL_FORM_PREV+'" AND PNA_POPULATION="'+GetField('PFF_POPULATION')+'") OR PFI_SALARIE="")'
     Else
          St := St + ' AND PFI_ETABLISSEMENT="'+GetField('PFF_ETABLISSEMENT')+'"';

     TobInsc:=Tob.Create('INSCFORMATION',Nil,-1);
     TobInsc.LoadDetailDBFromSQL('INSCFORMATION',St);

     If TobInsc.Detail.Count > 0 then
     begin
        St := 'Attention il existe des inscriptions au budget pour ';
        If (VH_Paie.PGForGestFraisByPop) Then
          St := St + 'cette population '
        Else
          St := St + 'cet établissement ';
        Rep:=PGIAsk(St + 'et ce lieu de formation,#13#10le données vont être mises a jour avec les nouvelles valeurs#13#10Voulez vous continuer ?',Ecran.Caption);
     //PT2 - Fin
        if rep <> mrYes then
        begin
           LastError:=1;  Exit;
        end;

        For i:=0 to TobInsc.Detail.Count - 1 Do
        Begin
            //PT2 - Début
            If (TobInsc.Detail[i].GetValue('PFI_SALARIE') = '') And (VH_Paie.PGForGestFraisByPop) Then
               If RecherchePopulation (TobInsc.Detail[i]) <> GetField('PFF_POPULATION') Then Continue;
            //PT2 - Fin

            FraisH:=GetField('PFF_FRAISHEBERG');
            FraisR:=GetField('PFF_FRAISREPAS');
            FraisT:=GetField('PFF_FRAISTRANSP');

            NbJour:=TobInsc.Detail[i].GetValue('PFI_JOURSTAGE');
            NbInsc:=TobInsc.Detail[i].GetValue('PFI_NBINSC');

            // Frais d'hébergement : 1 hébergement pour 2 jours, 2 pour 3, etc...
            If NbJour > 0 Then FraisH:=FraisH*(NbJour-1) Else FraisH := 0;  //PT2

            // Frais de repas : 1 repas / jour + 1 par hébergement
            If FraisH > 0 then FraisR := FraisR*((NbJour*2)-1) //PT1
            Else FraisR:=FraisR*NbJour; //PT2
            if FraisR < 0 then FraisR := 0;

            // Total : Hébergement + Repas + Transport
            TotalFrais:=FraisH+FraisR+FraisT;
            TobInsc.Detail[i].Putvalue('PFI_FRAISFORFAIT',TotalFrais*NbInsc);
            TobInsc.Detail[i].UpdateDB(False);
        End;
    end;

    FreeAndNil(TobInsc); //PT2
end;

procedure TOM_FORFAITFORM.OnAfterUpdateRecord ;
begin
end;

procedure TOM_FORFAITFORM.OnArgument ( S: String ) ;
var BDupliquer:TToolBarButton97;
begin
  Inherited ;
     Millesime:=Trim(ReadTokenPipe(S,';'));
     BDupliquer:=TToolBarButton97(GetControl('BDUPLIQUER'));
     If BDupliquer<>Nil then BDupliquer.OnClick:=BDupliquerClick;
     TFFiche(Ecran).DisabledMajCaption := True;

     //PT2 - Début
     // En cas de gestion par population, on affiche la combo et le libellé associé.
     // Sinon, on reste sur la sélection par établissement
     If (VH_Paie.PGForGestFraisByPop) Then
     Begin
          SetControlVisible('TPFF_POPULATION', True);
          SetControlVisible('PFF_POPULATION',  True);
          SetControlVisible('TPFF_ETABLISSEMENT', False);
          SetControlVisible('PFF_ETABLISSEMENT',  False);
          // Mise à jour de la liste de la combo avec les populations actives
          SetControlProperty('PFF_POPULATION', 'Plus', ' AND PPC_PREDEFINI="'+GetPredefiniPopulation(TYP_POPUL_FORM_PREV)+'"');
     End;
     //PT2 - Fin
end ;

procedure TOM_FORFAITFORM.BDupliquerClick(Sender:TObject);
var St:String;
begin
     St:=GetField('PFF_LIEUFORM')+';'+GetField('PFF_MILLESIME')+';'+IntToStr(getField('PFF_FRAISREPAS'))+';'+
         IntToStr(GetField('PFF_FRAISHEBERG'))+';'+IntToStr(GetField('PFF_FRAISTRANSP'));
     AglLanceFiche('PAY','MULETABFORFAIT', '', '',St);
end;

Initialization
  registerclasses ( [ TOM_FORFAITFORM ] ) ; 
end.
