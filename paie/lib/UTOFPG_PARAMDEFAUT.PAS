{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Unité de gestion de la saisie du paramétrage import ASC
Mots clefs ... : PAIE,PGIMPORT
*****************************************************************
PT1   : 24/07/2007 FC V_72 FQ 14423 Gestion des différents régimes
PT2-1 : 06/09/2007 VG V_80 Un modèle avec pour code '' n'est pas sauvegardé
                           FQ N°14664
PT2-2 : 06/09/2007 VG V_80 Suppression d'un modèle - FQ N°14664
}
unit UTOFPG_PARAMDEFAUT;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     LookUp,
     UTobDebug;

Type TMotif = record
     ControlMotif : TControl;    //TControl
     TitreMotif : string;        //Titre de la fenêtre
     TableMotif : string;        //Table
     ColonneMotif : string;      //Code
     SelectMotif : string;       //Libellé
     WhereMotif : string;        //Condition
     EditMotif : string;         //THEdit
     LabelMotif : string;        //TLabel
     end;

Type TClePGDefautParam = record
     Predefini : string;       //PDM_PREDEFINI
     Code : string;            //PDM_CODE
     LibelleParam : string;    //PDM_LIBELLEPARAM
     Defaut : string;          //PDM_DEFAUT
     end;

Type
  TOF_PG_PARAMDEFAUT = Class (TOF)
    public
      procedure OnNew                    ; override;
      procedure OnDelete                 ; override;
      procedure OnUpdate                 ; override;
      procedure OnLoad                   ; override;
      procedure OnArgument (S : String ) ; override;
      procedure OnDisplay                ; override;
      procedure OnClose                  ; override;
      procedure OnCancel                 ; override;
    private
      Modele : THEdit;
      LeMotif : TMotif;

      procedure ModeleCharge(TModeleEnCours : TOB);
      procedure ModeleChange(Sender: TObject);
      procedure ModeleElipsisClick (Sender: TObject);
      procedure EnregistreParamDefaut;
      procedure CreePGDefautParam (Cle : TClePGDefautParam; NomChamp,
                                   LibelleChamp, PgValeur : string; TDP : TOB);
      procedure MetABlanc;
      procedure MetABlancCle;
    end;

Implementation

procedure TOF_PG_PARAMDEFAUT.OnNew;
begin
Inherited;
MetABlancCle;
MetABlanc;
end;

procedure TOF_PG_PARAMDEFAUT.OnDelete;
var
StDefaut : string;
Rep : integer;
begin
Inherited;
//PT2-2
Rep:=PGIAsk ('Confirmez-vous la suppression du modèle en cours ?',
             'Suppression du modèle');
if Rep=mrNo then
   exit
else
   begin
   try
      BeginTrans;
      StDefaut:= 'DELETE'+
                 ' FROM PGDEFAUTPARAM WHERE'+
                 ' PDM_CODE="'+GetControlText ('MODELE')+'"';
      ExecuteSQL (StDefaut);
      MetABlancCle;
      MetABlanc;
      CommitTrans;
   Except
      Rollback;
      PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
              'Suppression du modèle');
      END;
   end;
//FIN PT2-2
end;

procedure TOF_PG_PARAMDEFAUT.OnUpdate;
begin
Inherited;
//PT2-1
if (GetControlText ('MODELE')<>'') then
   begin
   EnregistreParamDefaut;
   SetControlEnabled ('MODELE', False);
   SetControlEnabled ('LIBMODELE', False);
   end;
//FIN PT2-1
end;

procedure TOF_PG_PARAMDEFAUT.OnLoad;
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.OnArgument (S : String);
var
Lequel, State : string;
begin
Inherited;
State:= Trim (ReadTokenPipe (S,';'));
Lequel:= Trim (ReadTokenPipe (S,';')) ;

SetControlText ('MODELE', Lequel);

Modele:= THEdit (GetControl ('MODELE'));
if (Modele <> NIL) then
   begin
   Modele.OnElipsisClick:= ModeleElipsisClick;
   ModeleChange (Modele);
   end;
if (State<>'CREATION') then
   SetControlEnabled ('LIBMODELE', False);
end;

procedure TOF_PG_PARAMDEFAUT.OnClose;
begin
Inherited;
end;

procedure TOF_PG_PARAMDEFAUT.OnDisplay ();
begin
Inherited;
end;

procedure TOF_PG_PARAMDEFAUT.OnCancel ();
begin
Inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/05/2007
Modifié le ... :   /  /
Description .. : Chargement d'un modèle
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.ModeleCharge(TModeleEnCours : TOB);
var
TParam : TOB;
begin
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGRACINEAUXI'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGRACINEAUXI', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGNUMAUXI'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGNUMAUXI', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_HORAIREETABL'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_HORAIREETABL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_SMIC'], False);
if (TParam<>nil) then
   SetControlText ('ETB_SMIC', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PCTFRAISPROF'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_PCTFRAISPROF', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_TAUXVERSTRANS'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_TAUXVERSTRANS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_NBJOUTRAV'], False);
if (TParam<>nil) then
   SetControlText ('ETB_NBJOUTRAV', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_NBREACQUISCP'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_NBREACQUISCP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_NBACQUISCP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_NBACQUISCP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_NBRECPSUPP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_NBRECPSUPP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_VALANCCP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_VALANCCP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_DATEACQCPANC'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_DATEACQCPANC', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_VALODXMN'], False);
if (TParam<>nil) then
   SetControlText ('ETB_VALODXMN', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_UNITEPRISEFF'],
                                   False);
if (TParam<>nil) then
   SetControlText ('PSA_UNITEPRISEFF', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_SALAIRETHEO'],
                                   False);
if (TParam<>nil) then
   SetControlText ('PSA_SALAIRETHEO', TParam.GetValue ('PDM_PGVALEUR'));

//THValComboBox
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['CBACTIVITE'], False);
if (TParam<>nil) then
   SetControlText ('CBACTIVITE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGTYPENUMSAL'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGTYPENUMSAL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ET_JURIDIQUE'], False);
if (TParam<>nil) then
   SetControlText ('ET_JURIDIQUE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_TYPSMIC'], False);
if (TParam<>nil) then
   SetControlText ('ETB_TYPSMIC', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_STANDCALEND'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_STANDCALEND', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_JOURHEURE'], False);
if (TParam<>nil) then
   SetControlText ('ETB_JOURHEURE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_MEDTRAVGU'], False);
if (TParam<>nil) then
   SetControlText ('ETB_MEDTRAVGU', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_CODEDDTEFPGU'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_CODEDDTEFPGU', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILREM'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILREM', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PERIODBUL'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PERIODBUL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_REDREPAS'], False);
if (TParam<>nil) then
   SetControlText ('ETB_REDREPAS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_REDRTT1'], False);
if (TParam<>nil) then
   SetControlText ('ETB_REDRTT1', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILAFP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILAFP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILRET'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILRET', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILMUT'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILMUT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILANCIEN'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILANCIEN', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFIL'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFIL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILRBS'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILRBS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_REDRTT2'], False);
if (TParam<>nil) then
   SetControlText ('ETB_REDRTT2', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILPRE'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILPRE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILTSS'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILTSS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILAPP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILAPP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILTRANS'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILTRANS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PROFILCGE'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PROFILCGE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_1ERREPOSH'], False);
if (TParam<>nil) then
   SetControlText ('ETB_1ERREPOSH', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_EDITBULCP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_EDITBULCP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_RELIQUAT'], False);
if (TParam<>nil) then
   SetControlText ('ETB_RELIQUAT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_2EMEREPOSH'], False);
if (TParam<>nil) then
   SetControlText ('ETB_2EMEREPOSH', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_BASANCCP'], False);
if (TParam<>nil) then
   SetControlText ('ETB_BASANCCP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_TYPDATANC'], False);
if (TParam<>nil) then
   SetControlText ('ETB_TYPDATANC', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_VALORINDEMCP'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_VALORINDEMCP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_MVALOMS'], False);
if (TParam<>nil) then
   SetControlText ('ETB_MVALOMS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PAIEVALOMS'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PAIEVALOMS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_MOISPAIEMENT'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_MOISPAIEMENT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PGMODEREGLE'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_PGMODEREGLE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_BCMODEREGLE'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_BCMODEREGLE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_BCMOISPAIEMENT'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_BCMOISPAIEMENT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PAIACOMPTE'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PAIACOMPTE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PAIFRAIS'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PAIFRAIS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_CODESECTION'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_CODESECTION', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PRUDHCOLL'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PRUDHCOLL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PRUDHSECT'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PRUDHSECT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PRUDHVOTE'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PRUDHVOTE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PRUDH'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PRUDH', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_DADSSECTION'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_DADSSECTION', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_TYPDADSSECT'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_TYPDADSSECT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_ETATBULLETIN'],
                                   False);
if (TParam<>nil) then
   SetControlText ('PSA_ETATBULLETIN', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_PROFILTPS'], False);
if (TParam<>nil) then
   SetControlText ('PSA_PROFILTPS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_REGIMESS'], False);
if (TParam<>nil) then
   SetControlText ('PSA_REGIMESS', TParam.GetValue ('PDM_PGVALEUR'));
//DEB PT1
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_TYPEREGIME'], False);
if (TParam<>nil) then
   SetControlText ('PSA_TYPEREGIME', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_REGIMEMAL'], False);
if (TParam<>nil) then
   SetControlText ('PSA_REGIMEMAL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_REGIMEAT'], False);
if (TParam<>nil) then
   SetControlText ('PSA_REGIMEAT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_REGIMEVIP'], False);
if (TParam<>nil) then
   SetControlText ('PSA_REGIMEVIP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_REGIMEVIS'], False);
if (TParam<>nil) then
   SetControlText ('PSA_REGIMEVIS', TParam.GetValue ('PDM_PGVALEUR'));
//FIN PT1
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_DADSCAT'], False);
if (TParam<>nil) then
   SetControlText ('PSA_DADSCAT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_DADSPROF'], False);
if (TParam<>nil) then
   SetControlText ('PSA_DADSPROF', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_UNITETRAVAIL'],
                                   False);
if (TParam<>nil) then
   SetControlText ('PSA_UNITETRAVAIL', TParam.GetValue ('PDM_PGVALEUR'));

//TCheckBox
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGINCSALARIE'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGINCSALARIE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGTIERSAUXIAUTO'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGTIERSAUXIAUTO', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGDECALAGE'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGDECALAGE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGDECALAGEPETIT'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGDECALAGEPETIT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGANALYTIQUE'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGANALYTIQUE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGCONGES'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGCONGES', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGMODIFLIGNEIMP'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGMODIFLIGNEIMP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGABSENCE'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGABSENCE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGABSENCECHEVAL'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGABSENCECHEVAL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGCHR6SEMAINE'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGCHR6SEMAINE', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGMSA'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGMSA', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGINTERMITTENTS'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGINTERMITTENTS', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGBTP'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGBTP', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGRESPONSABLES'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGRESPONSABLES', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGCOEFFEVOSAL'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGCOEFFEVOSAL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGSAISIEARRET'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGSAISIEARRET', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGTICKETRESTAU'],
                                   False);
if (TParam<>nil) then
   SetControlText ('SO_PGTICKETRESTAU', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['SO_PGPCS2003'], False);
if (TParam<>nil) then
   SetControlText ('SO_PGPCS2003', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_SUBROGATION'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_SUBROGATION', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_PERIODCT'], False);
if (TParam<>nil) then
   SetControlText ('ETB_PERIODCT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_TRANS9SAL'], False);
if (TParam<>nil) then
   SetControlText ('ETB_TRANS9SAL', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_CONGESPAYES'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_CONGESPAYES', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['PSA_PRISEFFECTIF'],
                                   False);
if (TParam<>nil) then
   SetControlText ('PSA_PRISEFFECTIF', TParam.GetValue ('PDM_PGVALEUR'));

//THNumEdit
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_JOURPAIEMENT'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_JOURPAIEMENT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_BCJOURPAIEMENT'],
                                   False);
if (TParam<>nil) then
   SetControlText ('ETB_BCJOURPAIEMENT', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_JEDTDU'], False);
if (TParam<>nil) then
   SetControlText ('ETB_JEDTDU', TParam.GetValue ('PDM_PGVALEUR'));
TParam:= TModeleEnCours.FindFirst (['PDM_NOMCHAMP'], ['ETB_JEDTAU'], False);
if (TParam<>nil) then
   SetControlText ('ETB_JEDTAU', TParam.GetValue ('PDM_PGVALEUR'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Modification de la valeur du modèle
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.ModeleChange(Sender: TObject);
var
sWhere : string;
TFille, TModeleEnCours : TOB;
begin
MetABlanc;             //PT2-2
LeMotif.ControlMotif:= Modele;
LeMotif.TitreMotif:= 'Modèles d''import paie';
LeMotif.TableMotif:= 'PGDEFAUTPARAM';
LeMotif.ColonneMotif:= 'PDM_CODE';
LeMotif.SelectMotif:= 'PDM_LIBELLEPARAM';
LeMotif.WhereMotif:= '';
LeMotif.EditMotif:= 'MODELE';
LeMotif.LabelMotif:= 'LIBMODELE';

sWhere:= 'SELECT *'+
         ' FROM '+LeMotif.TableMotif+' WHERE '+
         LeMotif.ColonneMotif+' = "'+GetControlText (LeMotif.EditMotif)+'"';
TModeleEnCours:= TOB.Create ('Le modèle', Nil, -1);
TModeleEnCours.LoadDetailDBFromSQL (LeMotif.TableMotif, sWhere);
if (TModeleEnCours.FillesCount (1)<>0) then
   begin
   TFille:= TModeleEnCours.FindFirst ([''], [''], False);
   SetControlText (LeMotif.LabelMotif, TFille.GetValue ('PDM_LIBELLEPARAM'));
   SetControlText ('DEFAUT', TFille.GetValue ('PDM_DEFAUT'));
   end
else
   begin
   SetControlText(LeMotif.EditMotif, '');
   SetControlText(LeMotif.LabelMotif, '');
   SetControlText('DEFAUT', 'X');
   end;
ModeleCharge (TModeleEnCours);
FreeAndNil (TModeleEnCours);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Click sur l'ellipsis modèle d'import
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.ModeleElipsisClick(Sender: TObject);
var
StSelect, sWhere : string;
begin
ModeleChange (Sender);

StSelect:= 'SELECT '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif+
           ' FROM '+LeMotif.TableMotif;
if (LeMotif.WhereMotif<> '') then
   begin
   sWhere:= StSelect+' WHERE '+LeMotif.WhereMotif;
   sWhere:= sWhere+' GROUP BY '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif;
   end
else
   sWhere:= StSelect+' GROUP BY '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif;

{$IFNDEF EAGLCLIENT}
LookUpList (LeMotif.ControlMotif, LeMotif.TitreMotif, LeMotif.TableMotif,
            LeMotif.ColonneMotif, LeMotif.SelectMotif, LeMotif.WhereMotif,
            LeMotif.ColonneMotif, TRUE, -1, sWhere);
{$ELSE}
LookupList (LeMotif.ControlMotif, LeMotif.TitreMotif, LeMotif.TableMotif,
            LeMotif.ColonneMotif, LeMotif.SelectMotif, LeMotif.WhereMotif,
            LeMotif.ColonneMotif, TRUE, -1, StSelect);
{$ENDIF}
ModeleChange (Sender);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/05/2007
Modifié le ... :   /  /
Description .. : Enregistrement des données de la table PGDEFAUTPARAM
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.EnregistreParamDefaut;
var
TobChamp, TOB_PGDEFAUTPARAM, TChamp : TOB;
Cle : TClePGDefautParam;
S_Champ, StDefaut, StPrefixe : string;
begin
TOB_PGDEFAUTPARAM:= TOB.Create ('Mère PGDEFAUTPARAM', nil, -1);
//
StPrefixe:= '"ET", "ETB", "PSA"';
StDefaut:= 'SELECT PDM_CODE'+
           ' FROM PGDEFAUTPARAM WHERE'+
           ' PDM_CODE<>"'+GetControlText ('MODELE')+'"';
if (ExisteSQL (StDefaut)=True) then
   SetControlText ('DEFAUT', '-');

S_Champ:= 'SELECT DH_NOMCHAMP, DH_LIBELLE'+
             ' FROM DECHAMPS WHERE'+
             ' DH_PREFIXE IN ('+StPrefixe+')';
TobChamp := TOB.Create('Les Champs', NIL, -1);
TobChamp.LoadDetailDBFromSQL ('Les Champs', S_Champ);

if (FileExists('CEGID.par')) then
   Cle.Predefini:= 'CEG'
else
   Cle.Predefini:= 'STD';
Cle.Code:= GetControlText ('MODELE');
Cle.LibelleParam:= GetControlText ('LIBMODELE');
if (GetCheckBoxState ('DEFAUT')=cbChecked) then
   Cle.Defaut:= 'X'
else
   Cle.Defaut:= '-';
   
//THEdit
if (GetControlText ('SO_PGRACINEAUXI')<>'') then
   CreePGDefautParam (Cle, 'SO_PGRACINEAUXI', '',
                      GetControlText ('SO_PGRACINEAUXI'), TOB_PGDEFAUTPARAM);
if (GetControlText ('SO_PGNUMAUXI')<>'') then
   CreePGDefautParam (Cle, 'SO_PGNUMAUXI', '', GetControlText ('SO_PGNUMAUXI'),
                      TOB_PGDEFAUTPARAM);
if (GetControlText ('ETB_HORAIREETABL')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_HORAIREETABL'], TRUE);
   CreePGDefautParam (Cle, 'ETB_HORAIREETABL', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_HORAIREETABL'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_SMIC')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_SMIC'], TRUE);
   CreePGDefautParam (Cle, 'ETB_SMIC', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_SMIC'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PCTFRAISPROF')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PCTFRAISPROF'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PCTFRAISPROF', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PCTFRAISPROF'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_TAUXVERSTRANS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_TAUXVERSTRANS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_TAUXVERSTRANS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_TAUXVERSTRANS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_NBJOUTRAV')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_NBJOUTRAV'], TRUE);
   CreePGDefautParam (Cle, 'ETB_NBJOUTRAV', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_NBJOUTRAV'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_NBREACQUISCP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_NBREACQUISCP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_NBREACQUISCP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_NBREACQUISCP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_NBACQUISCP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_NBACQUISCP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_NBACQUISCP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_NBACQUISCP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_NBRECPSUPP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_NBRECPSUPP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_NBRECPSUPP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_NBRECPSUPP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_VALANCCP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_VALANCCP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_VALANCCP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_VALANCCP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_DATEACQCPANC')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_DATEACQCPANC'], TRUE);
   CreePGDefautParam (Cle, 'ETB_DATEACQCPANC', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_DATEACQCPANC'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_VALODXMN')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_VALODXMN'], TRUE);
   CreePGDefautParam (Cle, 'ETB_VALODXMN', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_VALODXMN'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_UNITEPRISEFF')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_UNITEPRISEFF'], TRUE);
   CreePGDefautParam (Cle, 'PSA_UNITEPRISEFF', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_UNITEPRISEFF'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_SALAIRETHEO')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_SALAIRETHEO'], TRUE);
   CreePGDefautParam (Cle, 'PSA_SALAIRETHEO', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_SALAIRETHEO'), TOB_PGDEFAUTPARAM);
   end;

//THValComboBox
if (GetControlText ('CBACTIVITE')<>'') then
   CreePGDefautParam (Cle, 'CBACTIVITE', '', GetControlText ('CBACTIVITE'),
                      TOB_PGDEFAUTPARAM);
if (GetControlText ('SO_PGTYPENUMSAL')<>'') then
   CreePGDefautParam (Cle, 'SO_PGTYPENUMSAL', '',
                      GetControlText ('SO_PGTYPENUMSAL'), TOB_PGDEFAUTPARAM);
if (GetControlText ('ET_JURIDIQUE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ET_JURIDIQUE'], TRUE);
   CreePGDefautParam (Cle, 'ET_JURIDIQUE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ET_JURIDIQUE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_TYPSMIC')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_TYPSMIC'], TRUE);
   CreePGDefautParam (Cle, 'ETB_TYPSMIC', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_TYPSMIC'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_STANDCALEND')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_STANDCALEND'], TRUE);
   CreePGDefautParam (Cle, 'ETB_STANDCALEND', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_STANDCALEND'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_JOURHEURE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_JOURHEURE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_JOURHEURE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_JOURHEURE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_MEDTRAVGU')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_MEDTRAVGU'], TRUE);
   CreePGDefautParam (Cle, 'ETB_MEDTRAVGU', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_MEDTRAVGU'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_CODEDDTEFPGU')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_CODEDDTEFPGU'], TRUE);
   CreePGDefautParam (Cle, 'ETB_CODEDDTEFPGU', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_CODEDDTEFPGU'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILREM')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILREM'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILREM', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILREM'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PERIODBUL')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PERIODBUL'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PERIODBUL', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PERIODBUL'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_REDREPAS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_REDREPAS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_REDREPAS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_REDREPAS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_REDRTT1')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_REDRTT1'], TRUE);
   CreePGDefautParam (Cle, 'ETB_REDRTT1', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_REDRTT1'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILAFP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILAFP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILAFP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILAFP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILRET')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILRET'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILRET', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILRET'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILMUT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILMUT'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILMUT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILMUT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILANCIEN')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILANCIEN'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILANCIEN', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILANCIEN'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFIL')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFIL'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFIL', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFIL'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILRBS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILRBS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILRBS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILRBS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_REDRTT2')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_REDRTT2'], TRUE);
   CreePGDefautParam (Cle, 'ETB_REDRTT2', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_REDRTT2'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILPRE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILPRE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILPRE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILPRE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILTSS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILTSS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILTSS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILTSS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILAPP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILAPP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILAPP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILAPP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILTRANS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILTRANS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILTRANS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILTRANS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PROFILCGE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PROFILCGE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PROFILCGE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PROFILCGE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_1ERREPOSH')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_1ERREPOSH'], TRUE);
   CreePGDefautParam (Cle, 'ETB_1ERREPOSH', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_1ERREPOSH'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_EDITBULCP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_EDITBULCP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_EDITBULCP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_EDITBULCP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_RELIQUAT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_RELIQUAT'], TRUE);
   CreePGDefautParam (Cle, 'ETB_RELIQUAT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_RELIQUAT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_2EMEREPOSH')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_2EMEREPOSH'], TRUE);
   CreePGDefautParam (Cle, 'ETB_2EMEREPOSH', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_2EMEREPOSH'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_BASANCCP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_BASANCCP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_BASANCCP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_BASANCCP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_TYPDATANC')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_TYPDATANC'], TRUE);
   CreePGDefautParam (Cle, 'ETB_TYPDATANC', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_TYPDATANC'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_VALORINDEMCP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_VALORINDEMCP'], TRUE);
   CreePGDefautParam (Cle, 'ETB_VALORINDEMCP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_VALORINDEMCP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_MVALOMS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_MVALOMS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_MVALOMS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_MVALOMS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PAIEVALOMS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PAIEVALOMS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PAIEVALOMS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PAIEVALOMS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_MOISPAIEMENT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_MOISPAIEMENT'], TRUE);
   CreePGDefautParam (Cle, 'ETB_MOISPAIEMENT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_MOISPAIEMENT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PGMODEREGLE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PGMODEREGLE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PGMODEREGLE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PGMODEREGLE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_BCMODEREGLE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_BCMODEREGLE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_BCMODEREGLE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_BCMODEREGLE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_BCMOISPAIEMENT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_BCMOISPAIEMENT'], TRUE);
   CreePGDefautParam (Cle, 'ETB_BCMOISPAIEMENT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_BCMOISPAIEMENT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PAIACOMPTE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PAIACOMPTE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PAIACOMPTE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PAIACOMPTE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PAIFRAIS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PAIFRAIS'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PAIFRAIS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PAIFRAIS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_CODESECTION')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_CODESECTION'], TRUE);
   CreePGDefautParam (Cle, 'ETB_CODESECTION', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_CODESECTION'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PRUDHCOLL')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PRUDHCOLL'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PRUDHCOLL', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PRUDHCOLL'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PRUDHSECT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PRUDHSECT'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PRUDHSECT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PRUDHSECT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PRUDHVOTE')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PRUDHVOTE'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PRUDHVOTE', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PRUDHVOTE'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_PRUDH')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PRUDH'], TRUE);
   CreePGDefautParam (Cle, 'ETB_PRUDH', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_PRUDH'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_DADSSECTION')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_DADSSECTION'], TRUE);
   CreePGDefautParam (Cle, 'ETB_DADSSECTION', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_DADSSECTION'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('ETB_TYPDADSSECT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_TYPDADSSECT'], TRUE);
   CreePGDefautParam (Cle, 'ETB_TYPDADSSECT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('ETB_TYPDADSSECT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_ETATBULLETIN')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_ETATBULLETIN'], TRUE);
   CreePGDefautParam (Cle, 'PSA_ETATBULLETIN', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_ETATBULLETIN'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_PROFILTPS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_PROFILTPS'], TRUE);
   CreePGDefautParam (Cle, 'PSA_PROFILTPS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_PROFILTPS'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_REGIMESS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_REGIMESS'], TRUE);
   CreePGDefautParam (Cle, 'PSA_REGIMESS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_REGIMESS'), TOB_PGDEFAUTPARAM);
   end;
//DEB PT1
if (GetControlText ('PSA_TYPEREGIME')<>'-') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_TYPEREGIME'], TRUE);
   CreePGDefautParam (Cle, 'PSA_TYPEREGIME', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_TYPEREGIME'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_REGIMEMAL')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_REGIMEMAL'], TRUE);
   CreePGDefautParam (Cle, 'PSA_REGIMEMAL', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_REGIMEMAL'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_REGIMEAT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_REGIMEAT'], TRUE);
   CreePGDefautParam (Cle, 'PSA_REGIMEAT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_REGIMEAT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_REGIMEVIP')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_REGIMEVIP'], TRUE);
   CreePGDefautParam (Cle, 'PSA_REGIMEVIP', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_REGIMEVIP'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_REGIMEVIS')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_REGIMEVIS'], TRUE);
   CreePGDefautParam (Cle, 'PSA_REGIMEVIS', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_REGIMEVIS'), TOB_PGDEFAUTPARAM);
   end;
//FIN PT1
if (GetControlText ('PSA_DADSCAT')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_DADSCAT'], TRUE);
   CreePGDefautParam (Cle, 'PSA_DADSCAT', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_DADSCAT'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_DADSPROF')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_DADSPROF'], TRUE);
   CreePGDefautParam (Cle, 'PSA_DADSPROF', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_DADSPROF'), TOB_PGDEFAUTPARAM);
   end;
if (GetControlText ('PSA_UNITETRAVAIL')<>'') then
   begin
   TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_UNITETRAVAIL'], TRUE);
   CreePGDefautParam (Cle, 'PSA_UNITETRAVAIL', TChamp.GetValue ('DH_LIBELLE'),
                      GetControlText ('PSA_UNITETRAVAIL'), TOB_PGDEFAUTPARAM);
   end;

//TCheckBox
CreePGDefautParam (Cle, 'SO_PGINCSALARIE', '',
                   GetControlText ('SO_PGINCSALARIE'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGTIERSAUXIAUTO', '',
                   GetControlText ('SO_PGTIERSAUXIAUTO'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGDECALAGE', '', GetControlText ('SO_PGDECALAGE'),
                   TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGDECALAGEPETIT', '',
                   GetControlText ('SO_PGDECALAGEPETIT'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGANALYTIQUE', '',
                   GetControlText ('SO_PGANALYTIQUE'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGCONGES', '', GetControlText ('SO_PGCONGES'),
                   TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGMODIFLIGNEIMP', '',
                   GetControlText ('SO_PGMODIFLIGNEIMP'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGABSENCE', '', GetControlText ('SO_PGABSENCE'),
                   TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGABSENCECHEVAL', '',
                   GetControlText ('SO_PGABSENCECHEVAL'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGCHR6SEMAINE', '',
                   GetControlText ('SO_PGCHR6SEMAINE'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGMSA', '', GetControlText ('SO_PGMSA'),
                   TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGINTERMITTENTS', '',
                   GetControlText ('SO_PGINTERMITTENTS'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGBTP', '', GetControlText ('SO_PGBTP'),
                   TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGRESPONSABLES', '',
                   GetControlText ('SO_PGRESPONSABLES'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGCOEFFEVOSAL', '',
                   GetControlText ('SO_PGCOEFFEVOSAL'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGSAISIEARRET', '',
                   GetControlText ('SO_PGSAISIEARRET'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGTICKETRESTAU', '',
                   GetControlText ('SO_PGTICKETRESTAU'), TOB_PGDEFAUTPARAM);
CreePGDefautParam (Cle, 'SO_PGPCS2003', '', GetControlText ('SO_PGPCS2003'),
                   TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_SUBROGATION'], TRUE);
CreePGDefautParam (Cle, 'ETB_SUBROGATION', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_SUBROGATION'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_PERIODCT'], TRUE);
CreePGDefautParam (Cle, 'ETB_PERIODCT', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_PERIODCT'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_TRANS9SAL'], TRUE);
CreePGDefautParam (Cle, 'ETB_TRANS9SAL', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_TRANS9SAL'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_CONGESPAYES'], TRUE);
CreePGDefautParam (Cle, 'ETB_CONGESPAYES', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_CONGESPAYES'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['PSA_PRISEFFECTIF'], TRUE);
CreePGDefautParam (Cle, 'PSA_PRISEFFECTIF', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('PSA_PRISEFFECTIF'), TOB_PGDEFAUTPARAM);

//THNumEdit
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_JOURPAIEMENT'], TRUE);
CreePGDefautParam (Cle, 'ETB_JOURPAIEMENT', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_JOURPAIEMENT'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_BCJOURPAIEMENT'], TRUE);
CreePGDefautParam (Cle, 'ETB_BCJOURPAIEMENT', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_BCJOURPAIEMENT'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_JEDTDU'], TRUE);
CreePGDefautParam (Cle, 'ETB_JEDTDU', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_JEDTDU'), TOB_PGDEFAUTPARAM);
TChamp:= TobChamp.FindFirst (['DH_NOMCHAMP'], ['ETB_JEDTAU'], TRUE);
CreePGDefautParam (Cle, 'ETB_JEDTAU', TChamp.GetValue ('DH_LIBELLE'),
                   GetControlText ('ETB_JEDTAU'), TOB_PGDEFAUTPARAM);

TOB_PGDEFAUTPARAM.SetAllModifie (TRUE);
TOB_PGDEFAUTPARAM.InsertOrUpdateDB (True);
FreeAndNil (TOB_PGDEFAUTPARAM);
FreeAndNil (TobChamp);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/05/2007
Modifié le ... :   /  /
Description .. : Création d'un élément PGDEFAUTPARAM dans la TOB associée
Mots clefs ... : PAIE,PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.CreePGDefautParam (Cle : TClePGDefautParam;
                                                NomChamp, LibelleChamp,
                                                PgValeur : string; TDP : TOB);
var
PGDefautParamFille : Tob;
begin
PGDefautParamFille:=TOB.Create ('PGDEFAUTPARAM', TDP, -1);
PGDefautParamFille.PutValue('PDM_PREDEFINI', Cle.Predefini);
PGDefautParamFille.PutValue('PDM_NODOSSIER', '000000');
PGDefautParamFille.PutValue('PDM_CODE', Cle.Code);
PGDefautParamFille.PutValue('PDM_LIBELLEPARAM', Cle.LibelleParam);
PGDefautParamFille.PutValue('PDM_DEFAUT', Cle.Defaut);

PGDefautParamFille.PutValue('PDM_NOMCHAMP', NomChamp);
PGDefautParamFille.PutValue('PDM_LIBELLE', LibelleChamp);
PGDefautParamFille.PutValue('PDM_PGVALEUR', PgValeur);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Réinitialisation de la fiche
Mots clefs ... : PAIE,PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.MetABlanc;
begin

//THEdit
SetControlText ('SO_PGRACINEAUXI', '');
SetControlText ('SO_PGNUMAUXI', '');
SetControlText ('ETB_HORAIREETABL', '');
SetControlText ('ETB_SMIC', '');
SetControlText ('ETB_PCTFRAISPROF', '');
SetControlText ('ETB_TAUXVERSTRANS', '');
SetControlText ('ETB_NBJOUTRAV', '');
SetControlText ('ETB_NBREACQUISCP', '');
SetControlText ('ETB_NBACQUISCP', '');
SetControlText ('ETB_NBRECPSUPP', '');
SetControlText ('ETB_VALANCCP', '');
SetControlText ('ETB_DATEACQCPANC', '');
SetControlText ('ETB_VALODXMN', '');
SetControlText ('PSA_UNITEPRISEFF', '');
SetControlText ('PSA_SALAIRETHEO', '');

//THValComboBox
SetControlText ('CBACTIVITE', '');
SetControlText ('SO_PGTYPENUMSAL', '');
SetControlText ('ET_JURIDIQUE', '');
SetControlText ('ETB_TYPSMIC', '');
SetControlText ('ETB_STANDCALEND', '');
SetControlText ('ETB_JOURHEURE', '');
SetControlText ('ETB_MEDTRAVGU', '');
SetControlText ('ETB_CODEDDTEFPGU', '');
SetControlText ('ETB_PROFILREM', '');
SetControlText ('ETB_PERIODBUL', '');
SetControlText ('ETB_REDREPAS', '');
SetControlText ('ETB_REDRTT1', '');
SetControlText ('ETB_PROFILAFP', '');
SetControlText ('ETB_PROFILRET', '');
SetControlText ('ETB_PROFILMUT', '');
SetControlText ('ETB_PROFILANCIEN', '');
SetControlText ('ETB_PROFIL', '');
SetControlText ('ETB_PROFILRBS', '');
SetControlText ('ETB_REDRTT2', '');
SetControlText ('ETB_PROFILPRE', '');
SetControlText ('ETB_PROFILTSS', '');
SetControlText ('ETB_PROFILAPP', '');
SetControlText ('ETB_PROFILTRANS', '');
SetControlText ('ETB_PROFILCGE', '');
SetControlText ('ETB_1ERREPOSH', '');
SetControlText ('ETB_EDITBULCP', '');
SetControlText ('ETB_RELIQUAT', '');
SetControlText ('ETB_2EMEREPOSH', '');
SetControlText ('ETB_BASANCCP', '');
SetControlText ('ETB_TYPDATANC', '');
SetControlText ('ETB_VALORINDEMCP', '');
SetControlText ('ETB_MVALOMS', '');
SetControlText ('ETB_PAIEVALOMS', '');
SetControlText ('ETB_MOISPAIEMENT', '');
SetControlText ('ETB_PGMODEREGLE', '');
SetControlText ('ETB_BCMODEREGLE', '');
SetControlText ('ETB_BCMOISPAIEMENT', '');
SetControlText ('ETB_PAIACOMPTE', '');
SetControlText ('ETB_PAIFRAIS', '');
SetControlText ('ETB_CODESECTION', '');
SetControlText ('ETB_PRUDHCOLL', '');
SetControlText ('ETB_PRUDHSECT', '');
SetControlText ('ETB_PRUDHVOTE', '');
SetControlText ('ETB_PRUDH', '');
SetControlText ('ETB_DADSSECTION', '');
SetControlText ('ETB_TYPDADSSECT', '');
SetControlText ('PSA_ETATBULLETIN', '');
SetControlText ('PSA_PROFILTPS', '');
SetControlText ('PSA_REGIMESS', '');
//DEB PT1
SetControlText ('PSA_REGIMEMAL', '');
SetControlText ('PSA_REGIMEAT', '');
SetControlText ('PSA_REGIMEVIP', '');
SetControlText ('PSA_REGIMEVIS', '');
//FIN PT1
SetControlText ('PSA_DADSCAT', '');
SetControlText ('PSA_DADSPROF', '');
SetControlText ('PSA_UNITETRAVAIL', '');

//TCheckBox
SetControlText ('DEFAUT', '-');
SetControlText ('SO_PGINCSALARIE', '-');
SetControlText ('SO_PGTIERSAUXIAUTO', '-');
SetControlText ('SO_PGDECALAGE', '-');
SetControlText ('SO_PGDECALAGEPETIT', '-');
SetControlText ('SO_PGANALYTIQUE', '-');
SetControlText ('SO_PGCONGES', '-');
SetControlText ('SO_PGMODIFLIGNEIMP', '-');
SetControlText ('SO_PGABSENCE', '-');
SetControlText ('SO_PGABSENCECHEVAL', '-');
SetControlText ('SO_PGCHR6SEMAINE', '-');
SetControlText ('SO_PGMSA', '-');
SetControlText ('SO_PGINTERMITTENTS', '-');
SetControlText ('SO_PGBTP', '-');
SetControlText ('SO_PGRESPONSABLES', '-');
SetControlText ('SO_PGCOEFFEVOSAL', '-');
SetControlText ('SO_PGSAISIEARRET', '-');
SetControlText ('SO_PGTICKETRESTAU', '-');
SetControlText ('SO_PGPCS2003', '-');
SetControlText ('ETB_SUBROGATION', '-');
SetControlText ('ETB_PERIODCT', '-');
SetControlText ('ETB_TRANS9SAL', '-');
SetControlText ('ETB_CONGESPAYES', '-');
SetControlText ('PSA_PRISEFFECTIF', '-');
SetControlText ('PSA_TYPEREGIME', '-');  //PT1

//THNumEdit
SetControlText ('ETB_JOURPAIEMENT', '1');
SetControlText ('ETB_BCJOURPAIEMENT', '1');
SetControlText ('ETB_JEDTDU', '1');
SetControlText ('ETB_JEDTAU', '31');

//THLabel
SetControlText ('TLIB_SMIC', '');
SetControlText ('NBACQUISCP_', '');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/09/2007
Modifié le ... :   /  /
Description .. : Réinitialisation de la clé de la fiche
Mots clefs ... : PAIE,PGIMPORT
*****************************************************************}
procedure TOF_PG_PARAMDEFAUT.MetABlancCle;
begin
SetControlEnabled ('MODELE', True);
SetControlEnabled ('LIBMODELE', True);

//THEdit
SetControlText ('MODELE', '');
SetControlText ('LIBMODELE', '');
end;


Initialization
registerclasses ([TOF_PG_PARAMDEFAUT]);
end.

