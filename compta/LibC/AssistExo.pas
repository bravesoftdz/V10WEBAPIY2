unit AssistExo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, HPanel, StdCtrls, Mask, Hctrls, HTB97, HSysMenu, hmsgbox,HEnt1,Ent1, ParamSoc;

type
  TFAssistExo = class(TForm)
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    Baide: TToolbarButton97;
    HPanel1: THPanel;
    HLabel1: THLabel;
    DATE_EXO: THCritMaskEdit;
    HLabel2: THLabel;
    HLabel3: THLabel;
    DATE_EXO_: THCritMaskEdit;
    Msg: THMsgBox;
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

procedure LanceAssistExo;
procedure CreationExerciceSurNouveauDossier ( NomBase : string; DateDebut, DateFin : TDateTime);

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibExercice, uTOB;   // CControleDureeExercice

{$R *.DFM}

procedure LanceAssistExo;
var  FAssistExo: TFAssistExo;
begin
  FAssistExo := TFAssistExo.Create (application);
  try
    FAssistExo.ShowModal;
  finally
    FAssistExo.Free;
  end;
end;

procedure TFAssistExo.BValiderClick(Sender: TObject);
Var D1,D2 : TDateTime ;
begin
  D1 := StrToDate(DATE_EXO.Text) ;
  D2 := StrToDate(DATE_EXO_.Text) ;

  if D2 < D1 then
  begin
    PGIInfo('La date de fin d''exercice doit être supérieure ou égale à la date de début.', 'Ouverture d''exercice');
    Date_Exo_.SetFocus;
    Exit;
  end;

  if not CControleDureeExercice( D1, D2) then
  begin
    PGIInfo('La durée d''exercice ne doit pas excéder 24 mois.', 'Ouverture d''exercice');
    Date_Exo_.SetFocus;
    Exit;
  end;

  CreationExerciceSurNouveauDossier ( '', D1,D2 );
  Msg.Execute(13,'','') ;
  RechargeParamSoc;
  VideTablettes;
  ModalResult := mrYes ;
end;

procedure CreationExerciceSurNouveauDossier ( NomBase : string; DateDebut, DateFin : TDateTime);
var
   ad,md,NbMois      :Word;
   DateExo           :TExoDate;
   DDeb              :TDateTime ;
   sLibelleExo       :string;
   stTableExercice   :string;                   
   strSQL            :string;

   strSociete        :string;   // $$$ JP 15/09/04 - code societe à calculer
   TOBSoc            :TOB;
begin
  DDeb:=0 ;
  DateExo.Deb := DateDebut;
  DateExo.Fin := DateFin;

  // $$$ JP: il faut prendre le code soc de la table DOSSIER de la base commune
  strSociete := '';
  if NomBase = '' then
  begin
       stTableExercice := 'EXERCICE';
       strSociete      := V_PGI.CodeSociete;
  end
  else
  begin
       stTableExercice := NomBase+'.dbo.EXERCICE';
       TOBSoc          := TOB.Create ('le code soc', nil, 1);
       try
          TOBSoc.LoadDetailFromSQL ('SELECT DOS_SOCIETE FROM DOSSIER WHERE DOS_NODOSSIER="' + Copy (NomBase, 3, Length (NomBase)-2) + '"');
          if TOBSoc.Detail.Count > 0 then
             strSociete := TOBSoc.Detail [0].GetString ('DOS_SOCIETE');
       finally
              TOBSoc.Free;
       end;
  end;

  NombreperExo(DateExo,md,ad,NbMois) ;
  if not CControleDureeExercice(DateDebut, DateFin) then exit;
  sLibelleExo := TraduireMemoire('du')+' '+DateToStr(DateDebut)+' '+TraduireMemoire('au')+' '+DateToStr(DateFin);
  strSQL := 'INSERT INTO '+stTableExercice+' (EX_EXERCICE,EX_LIBELLE,EX_ABREGE,EX_DATEDEBUT,EX_DATEFIN,' +
            'EX_ETATCPTA,EX_ETATBUDGET,EX_ETATADV,EX_ETATAPPRO,EX_ETATPROD,EX_SOCIETE,EX_VALIDEE,' +
            'EX_DATECUM,EX_DATECUMRUB,EX_DATECUMBUD,EX_DATECUMBUDGET,EX_BUDJAL,EX_NATEXO ) ' +
            'VALUES ("001","'+CheckdblQuote(sLibelleExo)+'","'+Copy(CheckdblQuote(sLibelleExo),1,14) +
            '","'+UsDateTime(DateDebut)+'",'+'"'+UsDateTime(DateFin)+'","OUV","OUV","NON","NON","NON","' +
            strSociete + '",'+'"------------------------","'+UsDateTime(DDeb)+'","'+UsDateTime(DDeb)+'","'+UsDateTime(2)+'","'+UsDateTime(2)+'","","")';

  ExecuteSql (strSQL);
  {'INSERT INTO '+stTableExercice+' (EX_EXERCICE,EX_LIBELLE,EX_ABREGE,EX_DATEDEBUT,EX_DATEFIN,'+
           'EX_ETATCPTA,EX_ETATBUDGET,EX_ETATADV,EX_ETATAPPRO,EX_ETATPROD,EX_SOCIETE,EX_VALIDEE,'+
           'EX_DATECUM,EX_DATECUMRUB,EX_DATECUMBUD,EX_DATECUMBUDGET,EX_BUDJAL,EX_NATEXO ) '+
           'VALUES ("001","'+CheckdblQuote(sLibelleExo)+'","'+Copy(CheckdblQuote(sLibelleExo),1,14)+'","'+UsDateTime(DateDebut)+'",'+
                    '"'+UsDateTime(DateFin)+'","OUV","OUV","NON","NON","NON","'+V_PGI.CodeSociete+'",'+
                    '"------------------------","'+UsDateTime(DDeb)+'","'+UsDateTime(DDeb)+'","'+UsDateTime(2)+'","'+UsDateTime(2)+'","","")') ;}
  if NomBase<> '' then ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="001" WHERE SOC_NOM="SO_CPEXOREF"')
  else  SetParamSoc('SO_CPEXOREF','001');
end;

end.
