import { useEffect, useMemo, useState } from "react"
import "./App.css"

const CAPICOM_CURRENT_USER_STORE = 2
const CAPICOM_MY_STORE = "My"
const CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = 2
const CADESCOM_CADES_BES = 1
const CADESCOM_BASE64_TO_BINARY = 1
const PLUGIN_WAIT_TIMEOUT_MS = 15000
const PLUGIN_WAIT_INTERVAL_MS = 250

function toBase64(arrayBuffer) {
  let binary = ""
  const bytes = new Uint8Array(arrayBuffer)
  const chunkSize = 0x8000
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize)
    binary += String.fromCharCode(...chunk)
  }
  return window.btoa(binary)
}

function getSubjectField(subjectName, fieldKeys) {
  if (!subjectName) {
    return ""
  }

  const normalized = subjectName.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map((part) => part.trim())
  for (const key of fieldKeys) {
    const match = normalized.find((part) => part.toUpperCase().startsWith(`${key.toUpperCase()}=`))
    if (match) {
      return match.slice(match.indexOf("=") + 1).replace(/^"|"$/g, "")
    }
  }

  return ""
}

function getCertificateOwnerInfo(subjectName) {
  const surname = getSubjectField(subjectName, ["SN", "SURNAME", "2.5.4.4"])
  const name = getSubjectField(subjectName, ["G", "GN", "GIVENNAME", "2.5.4.42"])
  const middleName = getSubjectField(subjectName, ["T", "INITIALS", "2.5.4.43", "2.5.4.41"])
  const cn = getSubjectField(subjectName, ["CN"])

  const fullNameByParts = [surname, name, middleName].filter(Boolean).join(" ")
  const fullName = fullNameByParts || cn || "Без ФИО"
  const email =
    getSubjectField(subjectName, ["E", "EMAILADDRESS", "1.2.840.113549.1.9.1"]) || "email не указан"

  return { fullName, email }
}

async function waitForCryptoPro() {
  const startedAt = Date.now()
  while (!window.cadesplugin && Date.now() - startedAt < PLUGIN_WAIT_TIMEOUT_MS) {
    await new Promise((resolve) => window.setTimeout(resolve, PLUGIN_WAIT_INTERVAL_MS))
  }

  if (!window.cadesplugin) {
    throw new Error("CryptoPro Browser Plugin не найден в браузере.")
  }

  let plugin = window.cadesplugin
  if (typeof plugin?.then === "function") {
    const timeoutPromise = new Promise((_, reject) => {
      window.setTimeout(() => {
        reject(new Error("Не удалось подключиться к CryptoPro Browser Plugin."))
      }, PLUGIN_WAIT_TIMEOUT_MS)
    })
    const readyPromise = new Promise((resolve, reject) => {
      plugin.then(
        () => resolve(true),
        (error) => reject(error || new Error("Плагин не инициализировался.")),
      )
    })

    await Promise.race([readyPromise, timeoutPromise])
  }

  if (!plugin || typeof plugin.CreateObjectAsync !== "function") {
    throw new Error(
      "Плагин установлен, но API CreateObjectAsync недоступен. Обычно это означает, что не поднят нативный компонент CryptoPro CSP/Extension Host или браузеру не даны права на localhost.",
    )
  }

  return { cadesplugin: plugin }
}

function getErrorMessage(error) {
  const message =
    (window.cadesplugin &&
      typeof window.cadesplugin.getLastError === "function" &&
      window.cadesplugin.getLastError(error)) ||
    error?.message ||
    String(error)

  return message || "Неизвестная ошибка"
}

function executeCades(cadesplugin, generatorFn) {
  const cadesTask = new Promise((resolve, reject) => {
    Promise.resolve()
      .then(() =>
        cadesplugin.async_spawn(function* () {
          const result = yield* generatorFn()
          resolve(result)
        }),
      )
      .catch(reject)
  })

  const timeoutTask = new Promise((_, reject) => {
    window.setTimeout(() => reject(new Error("Операция CryptoPro выполняется слишком долго.")), PLUGIN_WAIT_TIMEOUT_MS)
  })

  return Promise.race([cadesTask, timeoutTask])
}

function getEnvironmentSnapshot() {
  const pluginScript = document.querySelector('script[src*="cadesplugin_api.js"]')

  return [
    `URL: ${window.location.origin}`,
    `isSecureContext: ${String(window.isSecureContext)}`,
    `script cadesplugin_api.js: ${pluginScript ? "найден" : "не найден"}`,
    `window.cadesplugin: ${typeof window.cadesplugin}`,
  ]
}

function readCertificatesFromStore(cadesplugin) {
  return executeCades(cadesplugin, function* () {
    const store = yield cadesplugin.CreateObjectAsync("CAdESCOM.Store")
    yield store.Open(CAPICOM_CURRENT_USER_STORE, CAPICOM_MY_STORE, CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)

    const certs = yield store.Certificates
    const count = yield certs.Count
    const certList = []
    let certsWithPrivateKey = 0

    for (let i = 1; i <= count; i += 1) {
      const cert = yield certs.Item(i)
      const thumbprint = yield cert.Thumbprint
      const subjectName = yield cert.SubjectName
      const validFromDate = yield cert.ValidFromDate
      const validToDate = yield cert.ValidToDate
      const hasPrivateKey = yield cert.HasPrivateKey()
      const ownerInfo = getCertificateOwnerInfo(subjectName)

      if (hasPrivateKey) {
        certsWithPrivateKey += 1
      }

      certList.push({
        thumbprint,
        subjectName,
        hasPrivateKey,
        displayName: `ФИО: ${ownerInfo.fullName}; email: ${ownerInfo.email}; срок: с ${new Date(
          validFromDate,
        ).toLocaleDateString("ru-RU")} по ${new Date(validToDate).toLocaleDateString("ru-RU")})${
          hasPrivateKey ? "" : " - без закрытого ключа"
        }`,
      })
    }

    yield store.Close()

    return {
      openedStoreTitle: "CurrentUser/My",
      count,
      certList,
      certsWithPrivateKey,
    }
  })
}

function createDetachedSignature(cadesplugin, thumbprint, base64Data) {
  return executeCades(cadesplugin, function* () {
    const store = yield cadesplugin.CreateObjectAsync("CAdESCOM.Store")
    yield store.Open(CAPICOM_CURRENT_USER_STORE, CAPICOM_MY_STORE, CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)

    const certs = yield store.Certificates
    const found = yield certs.Find(0, thumbprint)
    const foundCount = yield found.Count

    if (foundCount < 1) {
      throw new Error("Сертификат недоступен. Повторно выберите сертификат.")
    }

    const cert = yield found.Item(1)
    const hasPrivateKey = yield cert.HasPrivateKey()
    if (!hasPrivateKey) {
      throw new Error("У выбранного сертификата нет закрытого ключа. Выберите сертификат с ключом.")
    }

    const signedData = yield cadesplugin.CreateObjectAsync("CAdESCOM.CadesSignedData")
    yield signedData.propset_ContentEncoding(CADESCOM_BASE64_TO_BINARY)
    yield signedData.propset_Content(base64Data)

    const signer = yield cadesplugin.CreateObjectAsync("CAdESCOM.CPSigner")
    yield signer.propset_Certificate(cert)

    const signed = yield signedData.SignCades(signer, CADESCOM_CADES_BES, true)
    yield store.Close()
    return signed
  })
}

export default function App() {
  const [certificates, setCertificates] = useState([])
  const [selectedThumbprint, setSelectedThumbprint] = useState("")
  const [status, setStatus] = useState("Ожидание проверки плагина")
  const [diagnostics, setDiagnostics] = useState([])
  const [documentFile, setDocumentFile] = useState(null)
  const [signature, setSignature] = useState("")
  const [isLoadingCertificates, setIsLoadingCertificates] = useState(false)
  const [isSigning, setIsSigning] = useState(false)
  const [environmentInfo, setEnvironmentInfo] = useState([])

  const selectedCertificateLabel = useMemo(() => {
    if (!selectedThumbprint) return "Сертификат не выбран"
    const selected = certificates.find((cert) => cert.thumbprint === selectedThumbprint)
    return selected ? selected.displayName : "Сертификат не выбран"
  }, [certificates, selectedThumbprint])

  function addDiagnostic(message) {
    const timestamp = new Date().toLocaleTimeString("ru-RU")
    setDiagnostics((prev) => [`[${timestamp}] ${message}`, ...prev].slice(0, 20))
  }

  function runEnvironmentCheck() {
    const lines = getEnvironmentSnapshot()

    if (!window.cadesplugin) {
      lines.push("Диагноз: расширение/плагин не инициализировал API на этой странице.")
      setEnvironmentInfo(lines)
      return
    }

    lines.push(`cadesplugin.CreateObjectAsync: ${typeof window.cadesplugin?.CreateObjectAsync}`)
    lines.push(`cadesplugin.then: ${typeof window.cadesplugin?.then}`)

    if (typeof window.cadesplugin?.CreateObjectAsync !== "function") {
      lines.push(
        "Диагноз: API объекта отсутствует. Обычно не запущен native host CryptoPro CSP или заблокирована связка расширения.",
      )
      setEnvironmentInfo(lines)
      return
    }

    waitForCryptoPro()
      .then(({ cadesplugin }) =>
        executeCades(cadesplugin, function* () {
          const about = yield cadesplugin.CreateObjectAsync("CAdESCOM.About")
          const cspName = yield about.CSPName(80)

          return {
            cspName: String(cspName),
          }
        }),
      )
      .then((checkResult) => {
        lines.push("Пробный вызов CreateObjectAsync: успешно.")
        lines.push(`CSPName: ${checkResult.cspName}`)
        lines.push("Диагноз: окружение готово, проблема не в инициализации API.")
      })
      .catch((error) => {
        lines.push(`Пробный вызов COM-объекта: ошибка: ${getErrorMessage(error)}`)
        lines.push("Диагноз: расширение есть, но нативный слой CryptoPro недоступен/не отвечает.")
      })
      .finally(() => setEnvironmentInfo(lines))
  }

  useEffect(() => {
    runEnvironmentCheck()
  }, [])

  function loadCertificates() {
    setIsLoadingCertificates(true)
    setStatus("Проверка CryptoPro Browser Plugin...")
    setSignature("")
    setDiagnostics([])
    addDiagnostic("Запрос списка сертификатов запущен.")

    addDiagnostic("Ожидание инициализации CryptoPro Browser Plugin.")

    waitForCryptoPro()
      .then(({ cadesplugin }) => {
        addDiagnostic("Плагин инициализирован.")
        return readCertificatesFromStore(cadesplugin)
      })
      .then((certData) => {
        const { openedStoreTitle, count, certList, certsWithPrivateKey } = certData
        addDiagnostic(`Открыто хранилище сертификатов: ${openedStoreTitle}.`)
        addDiagnostic(`Всего сертификатов в хранилище: ${count}.`)
        addDiagnostic("Хранилище сертификатов закрыто.")
        setCertificates(certList)

        if (certList.length === 0) {
          setSelectedThumbprint("")
          setStatus(`Сертификаты не найдены в хранилище ${openedStoreTitle}.`)
          addDiagnostic("Сертификаты не найдены.")
          return
        }

        const preferredCert = certList.find((cert) => cert.hasPrivateKey) ?? certList[0]
        setSelectedThumbprint(preferredCert.thumbprint)

        setStatus(
          `Найдено сертификатов: ${certList.length}. С закрытым ключом: ${certsWithPrivateKey}.`,
        )
        addDiagnostic(
          `Сертификаты загружены: ${certList.length}, с закрытым ключом: ${certsWithPrivateKey}.`,
        )
      })
      .catch((error) => {
        setCertificates([])
        setSelectedThumbprint("")
        const errorMessage = getErrorMessage(error)
        setStatus(`Ошибка загрузки сертификатов: ${errorMessage}`)
        addDiagnostic(`Ошибка загрузки сертификатов: ${errorMessage}`)
      })
      .finally(() => setIsLoadingCertificates(false))
  }

  function signFile() {
    if (!selectedThumbprint) {
      setStatus("Сначала выберите сертификат.")
      return
    }
    if (!documentFile) {
      setStatus("Сначала выберите файл для подписи.")
      return
    }

    setIsSigning(true)
    setStatus("Подписываю документ...")
    setSignature("")
    addDiagnostic(`Подписание файла "${documentFile.name}" запущено.`)

    waitForCryptoPro()
      .then(({ cadesplugin }) => {
        addDiagnostic("Плагин готов к подписанию.")

        return documentFile.arrayBuffer().then((buffer) => ({
          cadesplugin,
          base64Data: toBase64(buffer),
        }))
      })
      .then(({ cadesplugin, base64Data }) =>
        createDetachedSignature(cadesplugin, selectedThumbprint, base64Data),
      )
      .then((sign) => {
        setSignature(sign)
        setStatus("Файл успешно подписан.")
        addDiagnostic("Файл подписан успешно.")
      })
      .catch((error) => {
        const errorMessage = getErrorMessage(error)
        setStatus(`Ошибка подписи: ${errorMessage}`)
        addDiagnostic(`Ошибка подписи: ${errorMessage}`)
      })
      .finally(() => setIsSigning(false))
  }

  return (
    <main className="page">
      <header className="hero">
        <p className="hero-kicker">CryptoPro Workspace</p>
        <h1 className="page-title">Центр электронной подписи</h1>
        <p className="hero-subtitle">
          Выберите сертификат, загрузите документ и получите CAdES-BES подпись.
        </p>
      </header>

      <section className="card step-card">
        <h2 className="card-title">
          <span className="step-index">Шаг 1</span>
          <span>Сертификат</span>
        </h2>
        <div className="row">
          <button
            className="button"
            type="button"
            onClick={loadCertificates}
            disabled={isLoadingCertificates}
          >
            {isLoadingCertificates ? "Загрузка..." : "Выбрать сертификат"}
          </button>
          {certificates.length > 0 && (
            <select
              className="select"
              value={selectedThumbprint}
              onChange={(e) => setSelectedThumbprint(e.target.value)}
            >
              {certificates.map((cert) => (
                <option
                  key={cert.thumbprint}
                  value={cert.thumbprint}
                >
                  {cert.displayName}
                </option>
              ))}
            </select>
          )}
        </div>
        <p className="top-status">{status}</p>
        <p className="state">{selectedCertificateLabel}</p>
      </section>

      <section className="card step-card">
        <h2 className="card-title">
          <span className="step-index">Шаг 2</span>
          <span>Документ для подписи</span>
        </h2>
        <div className="row">
          <input
            className="native-file-input"
            type="file"
            onChange={(e) => setDocumentFile(e.target.files?.[0] ?? null)}
          />
          <span className="filename">{documentFile ? documentFile.name : "файл не выбран"}</span>
          <button
            className="button"
            type="button"
            onClick={signFile}
            disabled={isSigning}
          >
            {isSigning ? "Подписание..." : "Подписать файл"}
          </button>
        </div>
      </section>

      <section className="card step-card">
        <h2 className="card-title">
          <span className="step-index">Шаг 3</span>
          <span>Результат подписи</span>
        </h2>
        <textarea
          className="signature-output"
          value={signature}
          readOnly
          placeholder="Здесь будет отображена подпись"
        />
      </section>

      <section className="card">
        <h2 className="card-title">Текущий статус</h2>
        <p className="status">{status}</p>
      </section>

      <section className="card">
        <h2 className="card-title">Проверка окружения</h2>
        <div className="row">
          <button
            className="button"
            type="button"
            onClick={runEnvironmentCheck}
          >
            Перепроверить окружение
          </button>
        </div>
        <div className="diagnostics">
          {environmentInfo.length === 0 ? (
            <p className="status">Проверка окружения не выполнена.</p>
          ) : (
            environmentInfo.map((item, index) => (
              <p
                className="diagnostic-line"
                key={`${index}-${item}`}
              >
                {item}
              </p>
            ))
          )}
        </div>
      </section>

      <section className="card">
        <h2 className="card-title">Диагностика</h2>
        <div className="diagnostics">
          {diagnostics.length === 0 ? (
            <p className="status">Логи пока пустые. Нажмите "Выбрать сертификат".</p>
          ) : (
            diagnostics.map((item, index) => (
              <p
                className="diagnostic-line"
                key={`${index}-${item}`}
              >
                {item}
              </p>
            ))
          )}
        </div>
      </section>
    </main>
  )
}
