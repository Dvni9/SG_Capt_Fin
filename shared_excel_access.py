import os
import time
from contextlib import contextmanager


DEFAULT_LOCK_TIMEOUT = 30.0  # segundos
DEFAULT_LOCK_RETRY_DELAY = 0.5  # segundos
DEFAULT_STALE_LOCK_SECONDS = 5 * 60  # 5 minutos


class ExcelLockTimeout(Exception):
    pass


def _lock_path_for(excel_path: str) -> str:
    return excel_path + ".lock"


@contextmanager
def excel_file_lock(
    excel_path: str,
    timeout: float = DEFAULT_LOCK_TIMEOUT,
    retry_delay: float = DEFAULT_LOCK_RETRY_DELAY,
    stale_lock_seconds: float = DEFAULT_STALE_LOCK_SECONDS,
):
    lock_path = _lock_path_for(excel_path)
    start = time.time()

    while True:
        try:
            # Bloqueo cooperativo: crear un fichero .lock en la misma ruta compartida.
            # Si existe y es antiguo, lo consideramos "fantasma" (PC colgado) y lo limpiamos.
            try:
                st = os.stat(lock_path)
                age = time.time() - st.st_mtime
                if stale_lock_seconds is not None and age >= stale_lock_seconds:
                    os.remove(lock_path)
            except FileNotFoundError:
                pass

            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, f"locked_at={time.time()}\nexcel={excel_path}\n".encode("utf-8"))
            finally:
                os.close(fd)
            break
        except FileExistsError:
            elapsed = time.time() - start
            if elapsed >= timeout:
                raise ExcelLockTimeout(
                    f"No se pudo obtener el bloqueo para '{excel_path}' "
                    f"tras {int(timeout)} segundos. Otro equipo puede estar guardando."
                )
            time.sleep(retry_delay)

    try:
        yield
    finally:
        try:
            os.remove(lock_path)
        except FileNotFoundError:
            pass
        except PermissionError:
            # En shares de red a veces hay latencia; si no se puede borrar, no romper el flujo.
            pass
