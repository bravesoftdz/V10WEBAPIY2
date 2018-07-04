{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/08/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SALAIREFORM (SALAIREFORM)
Mots clefs ... : TOM;SALAIREFORM
*****************************************************************}
Unit UTomSalaireForm;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,
{$ELSE}
     eFiche,eFichList,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,EntPaie ;

Type
  TOM_SALAIREFORM = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
    Millesime:String;
    end ;

Implementation

procedure TOM_SALAIREFORM.OnNewRecord ;
begin
  Inherited ;
  SetField('PSF_MILLESIME',Millesime);
end ;

procedure TOM_SALAIREFORM.OnDeleteRecord ;
var Where:String;
begin
  Inherited ;
If VH_Paie.PGForValoSalaire='VCC' then Where:='WHERE PFI_LIBEMPLOIFOR="'+GetField('PSF_LIBEMPLOIFOR')+'" AND PFI_MILLESIME="'+GetField('PSF_MILLESIME')+'"'
Else Where:='WHERE PFI_SALARIE="" AND '+
'PFI_LIBEMPLOIFOR="'+GetField('PSF_LIBEMPLOIFOR')+'" AND PFI_MILLESIME="'+GetField('PSF_MILLESIME')+'"';
If ExisteSQl('SELECT PFI_CODESTAGE FROM INSCFORMATION '+Where ) Then
   begin
   PGIBox('Suppression impossible car ce libellé emploi est utilisé pour les inscriptions au budget de l''année '+GetField('PSF_MILLESIME'),Ecran.Caption);
   LastError:=1;
   Exit;
   end;
end ;

procedure TOM_SALAIREFORM.OnUpdateRecord ;
var Mess:String;
begin
  Inherited ;
Mess:='';
If GetField('PSF_LIBEMPLOIFOR')='' then Mess:='#13#10- Le libellé emploi';
If GetField('PSF_MILLESIME')='' then Mess:=Mess+'#13#10- Le millésime';
If Mess<>'' then
   begin
   PGIBox('Vous devez renseigné :'+Mess,Ecran.Caption);
   LastError:=1;
   end;
end ;

procedure TOM_SALAIREFORM.OnAfterUpdateRecord ;
var TobInsc,T,TobStage,TS:Tob;
    Q:TQuery;
    Duree:Integer;
    Salaire,NbInsc:Double;
    Where:String;
begin
  Inherited ;
Duree := 0;
If VH_Paie.PGForValoSalaire='VCC' then Where:='WHERE PFI_LIBEMPLOIFOR="'+GetField('PSF_LIBEMPLOIFOR')+'" AND PFI_MILLESIME="'+GetField('PSF_MILLESIME')+'"'
Else Where:='WHERE PFI_SALARIE="" AND '+
'PFI_LIBEMPLOIFOR="'+GetField('PSF_LIBEMPLOIFOR')+'" AND PFI_MILLESIME="'+GetField('PSF_MILLESIME')+'"';
If ExisteSQl('SELECT PFI_CODESTAGE FROM INSCFORMATION '+Where ) Then
   begin
   If PGIAsk('Voulez vous mettre à jour le budget','Salaires des participants')=MrYes Then
      begin
      Q:=OpenSQL('SELECT * FROM INSCFORMATION '+Where,True);
      TobInsc:=Tob.Create('INSCFORMATION',Nil,-1);
      TobInsc.LoadDetailDB('INSCFORMATION','','',Q,False);
      Ferme(Q);
      Q:=OpenSQL('SELECT PST_CODESTAGE,PST_DUREESTAGE FROM STAGE WHERE PST_MILLESIME="'+GetField('PSF_MILLESIME')+'"',True);
      TobStage:=Tob.Create('Les stages',Nil,-1);
      TobStage.LoadDetailDB('STAGE','','',Q,False);
      Ferme(Q);
      T:=TobInsc.FindFirst([''],[''],False);
      While T<>Nil do
            begin
            Salaire:=GetField('PSF_MONTANT');
            NbInsc:=T.GetValue('PFI_NBINSC');
            TS:=TobStage.FindFirst(['PST_CODESTAGE'],[T.GetValue('PFI_CODESTAGE')],False);
            If TS<>Nil Then Duree:=TS.GetValue('PST_DUREESTAGE');
            T.PutValue('PFI_COUTREELSAL',Salaire*NbInsc*Duree);
            T.UpdateDB;
            T:=TobInsc.FindNext([''],[''],False);
            end;
      TobInsc.Free;
      TobStage.Free;
      end;
   end;
end ;

procedure TOM_SALAIREFORM.OnLoadRecord ;
begin
  Inherited ;
TFFiche(Ecran).Caption:='Saisie des salaires catégoriels pour l''année '+Getfield('PSF_MILLESIME');
UpdateCaption(TFFiche(Ecran));
end ;

procedure TOM_SALAIREFORM.OnArgument ( S: String ) ;
begin
  Inherited ;
Millesime:=Trim(ReadTokenPipe(S,';'));
TFFiche(Ecran).DisabledMajCaption := True;
end ;

Initialization
  registerclasses ( [ TOM_SALAIREFORM ] ) ;
end.
