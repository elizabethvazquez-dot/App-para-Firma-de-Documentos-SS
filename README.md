/*******************************************************
 * SISTEMA DE FIRMA DE CARTAS DE SERVICIO SOCIAL
 * SERVICIO SOCIAL CAMPUS MONTERREY
 *******************************************************/

const CONFIG = {

  ORGANIZACION: 10,                  // J
  PROYECTO: 11,                      // K
  GRUPO: 13,                         // M
  NOMBRE_FIRMANTE: 22,               // V
  PUESTO: 23,                        // W
  CORREO: 24,                        // X

  USUARIO: 33,                       // AG
  CONTRASENA: 34,                    // AH

  LIGA_ACREDITACION: 42,             // AP
  ACREDITACION_FIRMADA: 43,           // AQ

  LIGA_COLABORACION: 44,             // AR
  COLABORACION_FIRMADA: 45,           // AS

  LIGA_CORRESPONSABILIDAD: 46,       // AT
  CORRESPONSABILIDAD_FIRMADA: 47,     // AU

};


/*******************************************************
 * CONFIGURACIÓN DE HOJA
 *******************************************************/

const SHEET_NAME = "Hoja 1";


/*******************************************************
 * OBTENER HOJA
 *******************************************************/

function getSheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      SHEET_NAME
    );

  if (!sheet) {

    throw new Error(
      "No se encontró la hoja: " +
      SHEET_NAME
    );

  }

  return sheet;

}


/*******************************************************
 * ABRIR APLICACIÓN
 *******************************************************/

function doGet() {

  return HtmlService

    .createHtmlOutputFromFile(
      "Index"
    )

    .setTitle(
      "Servicio Social Campus Monterrey"
    )

    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/*******************************************************
 * NORMALIZAR VALOR
 *******************************************************/

function normalizarValor_(valor) {

  return String(
    valor === null ||
    valor === undefined
      ? ""
      : valor
  )
    .trim();

}


/*******************************************************
 * LOGIN
 *******************************************************/

function login(
  usuario,
  password
) {

  try {

    const sheet =
      getSheet_();

    const lastRow =
      sheet.getLastRow();

    const lastColumn =
      Math.max(
        sheet.getLastColumn(),
        CONFIG.LIGA_CORRESPONSABILIDAD
      );

    if (lastRow < 2) {

      return {

        ok: false,

        mensaje:
          "No hay información registrada."

      };

    }

    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          lastColumn
        )
        .getDisplayValues();

    const proyectos = [];

    const usuarioBuscado =
      normalizarValor_(
        usuario
      );

    const passwordBuscada =
      normalizarValor_(
        password
      );

    for (
      let i = 0;
      i < data.length;
      i++
    ) {

      const row =
        data[i];

      const usuarioHoja =
        normalizarValor_(
          row[
            CONFIG.USUARIO - 1
          ]
        );

      const passwordHoja =
        normalizarValor_(
          row[
            CONFIG.CONTRASENA - 1
          ]
        );

      if (
        usuarioHoja !==
        usuarioBuscado
      ) {

        continue;

      }

      if (
        passwordHoja !==
        passwordBuscada
      ) {

        continue;

      }

      const ligaAcreditacion =
        normalizarValor_(
          row[
            CONFIG.LIGA_ACREDITACION - 1
          ]
        );

      const ligaColaboracion =
        normalizarValor_(
          row[
            CONFIG.LIGA_COLABORACION - 1
          ]
        );

      const ligaCorresponsabilidad =
        normalizarValor_(
          row[
            CONFIG.LIGA_CORRESPONSABILIDAD - 1
          ]
        );

      const acreditacionFirmada =
        normalizarValor_(
          row[
            CONFIG.ACREDITACION_FIRMADA - 1
          ]
        )
          .toUpperCase() === "SI";

      const colaboracionFirmada =
        normalizarValor_(
          row[
            CONFIG.COLABORACION_FIRMADA - 1
          ]
        )
          .toUpperCase() === "SI";

      const corresponsabilidadFirmada =
        normalizarValor_(
          row[
            CONFIG.CORRESPONSABILIDAD_FIRMADA - 1
          ]
        )
          .toUpperCase() === "SI";

      proyectos.push({

        rowNumber:
          i + 2,

        organization:
          row[
            CONFIG.ORGANIZACION - 1
          ] || "",

        project:
          row[
            CONFIG.PROYECTO - 1
          ] || "",

        group:
          row[
            CONFIG.GRUPO - 1
          ] || "",

        signer:
          row[
            CONFIG.NOMBRE_FIRMANTE - 1
          ] || "",

        position:
          row[
            CONFIG.PUESTO - 1
          ] || "",

        email:
          row[
            CONFIG.CORREO - 1
          ] || "",

        ligaAcreditacion:
          ligaAcreditacion,

        acreditacionFirmada:
          acreditacionFirmada,

        ligaColaboracion:
          ligaColaboracion,

        colaboracionFirmada:
          colaboracionFirmada,

        ligaCorresponsabilidad:
          ligaCorresponsabilidad,

        corresponsabilidadFirmada:
          corresponsabilidadFirmada

      });

    }

    if (
      proyectos.length === 0
    ) {

      return {

        ok: false,

        mensaje:
          "Usuario o contraseña incorrectos."

      };

    }

    return {

      ok: true,

      proyectos:
        proyectos

    };

  }

  catch (error) {

    return {

      ok: false,

      mensaje:
        "Error al iniciar sesión: " +
        error.message

    };

  }

}


/*******************************************************
 * OBTENER PDF
 *******************************************************/

function obtenerPDF(
  liga
) {

  try {

    const fileId =
      extraerIdDrive_(
        liga
      );

    if (!fileId) {

      return {

        ok: false,

        mensaje:
          "No se pudo obtener el ID del archivo PDF."

      };

    }

    const file =
      DriveApp.getFileById(
        fileId
      );

    const blob =
      file.getBlob();

    const bytes =
      blob.getBytes();

    return {

      ok: true,

      base64:
        Utilities.base64Encode(
          bytes
        ),

      fileId:
        fileId,

      nombre:
        file.getName()

    };

  }

  catch (error) {

    return {

      ok: false,

      mensaje:
        error.message

    };

  }

}


/*******************************************************
 * DESCARGAR PDF
 *******************************************************/

function descargarPDF(
  liga
) {

  try {

    if (!liga) {

      return {

        ok: false,

        mensaje:
          "No se recibió la liga de la carta."

      };

    }

    const fileId =
      extraerIdDrive_(
        liga
      );

    if (!fileId) {

      return {

        ok: false,

        mensaje:
          "No se pudo obtener el ID del archivo."

      };

    }

    const file =
      DriveApp.getFileById(
        fileId
      );

    const blob =
      file.getBlob();

    const bytes =
      blob.getBytes();

    let nombre =
      file.getName();

    if (
      !nombre
        .toLowerCase()
        .endsWith(".pdf")
    ) {

      nombre += ".pdf";

    }

    return {

      ok: true,

      base64:
        Utilities.base64Encode(
          bytes
        ),

      nombre:
        nombre,

      mimeType:
        "application/pdf"

    };

  }

  catch (error) {

    return {

      ok: false,

      mensaje:
        "No fue posible descargar la carta: " +
        error.message

    };

  }

}


/*******************************************************
 * GUARDAR PDF FIRMADO
 *******************************************************/

function guardarPDFFirmado(
  rowNumber,
  fileId,
  pdfBase64,
  tipoCarta
) {

  try {

    if (!fileId) {

      throw new Error(
        "No se recibió el ID del archivo PDF."
      );

    }

    if (!pdfBase64) {

      throw new Error(
        "No se recibió el PDF firmado."
      );

    }

    DriveApp.getFileById(
      fileId
    );

    const pdfBytes =
      Utilities.base64Decode(
        pdfBase64
      );

    const token =
      ScriptApp.getOAuthToken();

    const url =
      "https://www.googleapis.com/upload/drive/v3/files/" +
      encodeURIComponent(
        fileId
      ) +
      "?uploadType=media";

    const response =
      UrlFetchApp.fetch(
        url,
        {

          method:
            "patch",

          contentType:
            "application/pdf",

          payload:
            pdfBytes,

          headers: {

            Authorization:
              "Bearer " +
              token

          },

          muteHttpExceptions:
            true

        }
      );

    const responseCode =
      response.getResponseCode();

    if (
      responseCode < 200 ||
      responseCode >= 300
    ) {

      const detalle =
        response.getContentText();

      throw new Error(
        "Google Drive no pudo actualizar el archivo. Código: " +
        responseCode +
        " - " +
        detalle
      );

    }

    const sheet =
      getSheet_();

    const tipo =
      String(
        tipoCarta || ""
      )
        .trim()
        .toLowerCase();

    if (
      tipo ===
      "acreditacion"
    ) {

      sheet
        .getRange(
          rowNumber,
          CONFIG.ACREDITACION_FIRMADA
        )
        .setValue(
          "SI"
        );

    }

    else if (
      tipo ===
      "colaboracion"
    ) {

      sheet
        .getRange(
          rowNumber,
          CONFIG.COLABORACION_FIRMADA
        )
        .setValue(
          "SI"
        );

    }

    else if (
      tipo ===
      "corresponsabilidad"
    ) {

      sheet
        .getRange(
          rowNumber,
          CONFIG.CORRESPONSABILIDAD_FIRMADA
        )
        .setValue(
          "SI"
        );

    }

    SpreadsheetApp.flush();

    return {

      ok: true,

      fileId:
        fileId,

      tipoCarta:
        tipo,

      mensaje:
        "PDF actualizado correctamente."

    };

  }

  catch (error) {

    return {

      ok: false,

      mensaje:
        error.message

    };

  }

}


/*******************************************************
 * EXTRAER ID DE DRIVE
 *******************************************************/

function extraerIdDrive_(
  liga
) {

  if (!liga) {

    return null;

  }

  const texto =
    String(
      liga
    ).trim();

  let match =
    texto.match(
      /\/file\/d\/([a-zA-Z0-9_-]+)/
    );

  if (match) {

    return match[1];

  }

  match =
    texto.match(
      /[?&]id=([a-zA-Z0-9_-]+)/
    );

  if (match) {

    return match[1];

  }

  match =
    texto.match(
      /\/d\/([a-zA-Z0-9_-]+)/
    );

  if (match) {

    return match[1];

  }

  if (
    /^[a-zA-Z0-9_-]+$/.test(
      texto
    )
  ) {

    return texto;

  }

  return null;

}


/*******************************************************
 * AUTORIZAR GOOGLE DRIVE
 *******************************************************/

function autorizarDrive() {

  const root =
    DriveApp.getRootFolder();

  const token =
    ScriptApp.getOAuthToken();

  if (!token) {

    throw new Error(
      "No fue posible obtener el token OAuth."
    );

  }

  Logger.log(
    "Google Drive autorizado correctamente."
  );

  Logger.log(
    "Carpeta raíz: " +
    root.getName()
  );

  return (
    "Google Drive autorizado correctamente."
  );

}


/*******************************************************
 * PRUEBA DEL SISTEMA
 *******************************************************/

function pruebaSistema() {

  const sheet =
    getSheet_();

  return {

    hoja:
      sheet.getName(),

    filas:
      sheet.getLastRow(),

    columnas:
      sheet.getLastColumn(),

    usuario:
      CONFIG.USUARIO,

    contrasena:
      CONFIG.CONTRASENA,

    acreditacion:
      CONFIG.LIGA_ACREDITACION,

    acreditacionFirmada:
      CONFIG.ACREDITACION_FIRMADA,

    colaboracion:
      CONFIG.LIGA_COLABORACION,

    colaboracionFirmada:
      CONFIG.COLABORACION_FIRMADA,

    corresponsabilidad:
      CONFIG.LIGA_CORRESPONSABILIDAD,

    corresponsabilidadFirmada:
      CONFIG.CORRESPONSABILIDAD_FIRMADA

  };

}

