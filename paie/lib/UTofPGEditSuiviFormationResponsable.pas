{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSUIVIFORRESP ()
Mots clefs ... : TOF;PGSUIVIFORRESP
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
}
Unit UTofPGEditSuiviFormationResponsable;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
{$ELSE}
     eQRS1,UTOB,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,PGOutilsFormation ;

Type
  TOF_PGSUIVIFORRESP = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGSUIVIFORRESP.OnUpdate ;
var Millesime:THValComboBox;
    DateDebut,DateFin:TDateTime;
    Q:TQuery;
    SQL:String;
begin
  Inherited ;
Millesime:=THValComboBox(GetControl('MILLESIME'));
Q:=OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+Millesime.Value+'"',True);
Datedebut:=date;
DateFin:=Date;
if not Q.eof then
   begin
   DateDebut:=Q.FindField('PFE_DATEDEBUT').AsDateTime;
   DateFin:=Q.FindField('PFE_DATEFIN').AsDateTime;
   end;
Ferme(Q);
SetControlText('DATEDEBUT',DateToStr(DateDebut));
SetControlText('DATEFIN',DateToStr(DateFin));
SQL:='SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES '+
     'WHERE (EXISTS (SELECT PFO_RESPONSFOR FROM FORMATIONS WHERE PFO_RESPONSFOR=PSA_SALARIE AND '+
     'PFO_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PFO_DATEFIN<="'+UsDateTime(DateFin)+'") '+
     'OR EXISTS (SELECT PFI_RESPONSFOR FROM INSCFORMATION WHERE PFI_RESPONSFOR=PSA_SALARIE AND '+
     'PFI_MILLESIME="'+Millesime.Value+'")) '+
      'ORDER BY PSA_SALARIE';
SetcontrolText('XX_RUPTURE1','PSA_SALARIE');
TFQRS1(Ecran).WhereSQL:=SQL;
end ;

procedure TOF_PGSUIVIFORRESP.OnArgument (S : String ) ;
var Millesime : String;
    DD,DF : TDateTime;
    Num:Integer;
begin
  Inherited ;
   Millesime := RendMillesimeRealise(DD,DF);
   SetControlProperty('MILLESIME','Value',Millesime);
   SetControltext('DATEDEBUT',DateToStr(DD));
   SetControltext('DATEFIN',DateToStr(DF));
   For Num := 1 to VH_Paie.NBFormationLibre do
    begin
    if Num >8 then Break;
    VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
    end;
SetControlVIsible('PFORMATION',True);
end ;


Initialization
  registerclasses ( [ TOF_PGSUIVIFORRESP ] ) ;
end.
