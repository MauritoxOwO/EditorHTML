// Recurso web independiente del PCF. Registrar onLoad y onSave en el formulario.
(function (global) {
  "use strict";

  const TABLE_NAME = "ays_versionanuncio";
  const DIRTY_COLUMN = "ays_documentdirty_sn";
  const NOTIFICATION_ID = "ays_documentdirty_guard";
  const AUTO_SAVE_MODE = 70;
  const openDialogs = new WeakSet();
  const DIRTY_MESSAGE =
    "El documento contiene cambios sin guardar. Usa el botón Guardar del editor " +
    "antes de guardar el formulario o salir. Guardar el formulario no guarda el documento.";

  function getTargetForm(executionContext) {
    if (!executionContext || typeof executionContext.getFormContext !== "function") {
      global.console.error("[DocumentDirtyGuard] Falta el contexto de ejecución del formulario.");
      return null;
    }
    const formContext = executionContext.getFormContext();
    if (formContext.data.entity.getEntityName() !== TABLE_NAME) return null;
    return formContext;
  }

  function readState(formContext) {
    const attribute = formContext.getAttribute(DIRTY_COLUMN);
    if (!attribute) {
      return {
        attribute: null,
        blocked: true,
        level: "ERROR",
        title: "Revisar configuración del editor",
        message: "No se encuentra la columna " + DIRTY_COLUMN +
          " en este formulario. Añádela y vincúlala a documentDirty del PCF antes de continuar.",
      };
    }
    return {
      attribute: attribute,
      blocked: attribute.getValue() === true,
      level: "WARNING",
      title: "Documento sin guardar",
      message: DIRTY_MESSAGE,
    };
  }

  function updateNotification(formContext, state) {
    if (state.blocked) {
      formContext.ui.setFormNotification(state.message, state.level, NOTIFICATION_ID);
    } else {
      // No se eliminan avisos de otras bibliotecas del formulario.
      formContext.ui.clearFormNotification(NOTIFICATION_ID);
    }
  }

  function showAlert(formContext, state) {
    const key = state.attribute || formContext.data.entity;
    if (openDialogs.has(key)) return;
    openDialogs.add(key);

    function closeDialog() {
      openDialogs.delete(key);
    }
    function reportError(error) {
      closeDialog();
      global.console.error("[DocumentDirtyGuard] No se pudo abrir el aviso:", error);
    }

    try {
      global.Xrm.Navigation.openAlertDialog({
        title: state.title,
        text: state.message,
        confirmButtonLabel: "Entendido",
      }, { width: 500 }).then(closeDialog, reportError);
    } catch (error) {
      // El guardado ya está cancelado; sigue visible el aviso del formulario.
      reportError(error);
    }
  }

  function onDirtyChanged(executionContext) {
    const formContext = getTargetForm(executionContext);
    if (!formContext) return;
    updateNotification(formContext, readState(formContext));
  }

  function onLoad(executionContext) {
    const formContext = getTargetForm(executionContext);
    if (!formContext) return;
    const state = readState(formContext);
    if (state.attribute) {
      // OnLoad puede repetirse: registrar una única suscripción propia.
      state.attribute.removeOnChange(onDirtyChanged);
      state.attribute.addOnChange(onDirtyChanged);
    }
    const process = formContext.data.process;
    if (process && typeof process.addOnPreStageChange === "function" &&
        typeof process.removeOnPreStageChange === "function") {
      process.removeOnPreStageChange(onPreStageChange);
      process.addOnPreStageChange(onPreStageChange);
    }
    updateNotification(formContext, state);
  }

  function onSave(executionContext) {
    const formContext = getTargetForm(executionContext);
    if (!formContext) return;
    // Leer el atributo ahora; no depender de que OnChange ya haya actualizado el aviso.
    const state = readState(formContext);
    const eventArgs = executionContext.getEventArgs();
    if (state.blocked) {
      // Debe ejecutarse sincrónicamente, antes de abrir cualquier diálogo.
      // Se cancelan TODOS los modos para impedir que se guarde el indicador en Sí.
      eventArgs.preventDefault();
    }
    updateNotification(formContext, state);
    if (state.blocked && eventArgs.getSaveMode() !== AUTO_SAVE_MODE) {
      showAlert(formContext, state);
    }
  }

  function onPreStageChange(executionContext) {
    const formContext = getTargetForm(executionContext);
    if (!formContext) return;
    const state = readState(formContext);
    if (state.blocked) executionContext.getEventArgs().preventDefault();
    updateNotification(formContext, state);
    if (state.blocked) showAlert(formContext, state);
  }

  global.AYS = global.AYS || {};
  global.AYS.DocumentDirtyGuard = {
    onLoad: onLoad,
    onSave: onSave,
  };
})(window);
