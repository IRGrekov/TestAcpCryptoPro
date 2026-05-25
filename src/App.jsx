import { useMemo, useState } from "react"
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

function base64ToBytes(base64) {
  const binary = window.atob(base64.replace(/\s/g, ""))
  const bytes = new Uint8Array(binary.length)

  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i)
  }

  return bytes
}

function getSignatureFileName(fileName) {
  return `${fileName}.sig`
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement("a")

  link.href = url
  link.download = fileName
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
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
        (error) => reject(error || new Error("Плагин не инициализировался."))
      )
    })

    await Promise.race([readyPromise, timeoutPromise])
  }

  if (!plugin || typeof plugin.CreateObjectAsync !== "function") {
    throw new Error(
      "Плагин установлен, но API CreateObjectAsync недоступен. Обычно это означает, что не поднят нативный компонент CryptoPro CSP/Extension Host или браузеру не даны права на localhost."
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
        })
      )
      .catch(reject)
  })

  const timeoutTask = new Promise((_, reject) => {
    window.setTimeout(
      () => reject(new Error("Операция CryptoPro выполняется слишком долго.")),
      PLUGIN_WAIT_TIMEOUT_MS
    )
  })

  return Promise.race([cadesTask, timeoutTask])
}

function readCertificatesFromStore(cadesplugin) {
  return executeCades(cadesplugin, function* () {
    const store = yield cadesplugin.CreateObjectAsync("CAdESCOM.Store")
    yield store.Open(
      CAPICOM_CURRENT_USER_STORE,
      CAPICOM_MY_STORE,
      CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED
    )

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
          validFromDate
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

function createCadesDetachedSignature(cadesplugin, thumbprint, base64Data) {
  return executeCades(cadesplugin, function* () {
    const store = yield cadesplugin.CreateObjectAsync("CAdESCOM.Store")
    yield store.Open(
      CAPICOM_CURRENT_USER_STORE,
      CAPICOM_MY_STORE,
      CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED
    )

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
  const [documentFile, setDocumentFile] = useState(null)
  const [signature, setSignature] = useState("")
  const [isLoadingCertificates, setIsLoadingCertificates] = useState(false)
  const [isSigning, setIsSigning] = useState(false)

  const selectedCertificateLabel = useMemo(() => {
    if (!selectedThumbprint) return "Сертификат не выбран"
    const selected = certificates.find((cert) => cert.thumbprint === selectedThumbprint)
    return selected ? selected.displayName : "Сертификат не выбран"
  }, [certificates, selectedThumbprint])

  function loadCertificates() {
    setIsLoadingCertificates(true)
    setStatus("Проверка CryptoPro Browser Plugin...")
    setSignature("")

    waitForCryptoPro()
      .then(({ cadesplugin }) => readCertificatesFromStore(cadesplugin))
      .then((certData) => {
        const { openedStoreTitle, certList, certsWithPrivateKey } = certData
        setCertificates(certList)

        if (certList.length === 0) {
          setSelectedThumbprint("")
          setStatus(`Сертификаты не найдены в хранилище ${openedStoreTitle}.`)
          return
        }

        const preferredCert = certList.find((cert) => cert.hasPrivateKey) ?? certList[0]
        setSelectedThumbprint(preferredCert.thumbprint)

        setStatus(
          `Найдено сертификатов: ${certList.length}. С закрытым ключом: ${certsWithPrivateKey}.`
        )
      })
      .catch((error) => {
        setCertificates([])
        setSelectedThumbprint("")
        const errorMessage = getErrorMessage(error)
        setStatus(`Ошибка загрузки сертификатов: ${errorMessage}`)
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
    setStatus("Формирую SIG-подпись...")
    setSignature("")

    waitForCryptoPro()
      .then(({ cadesplugin }) =>
        documentFile.arrayBuffer().then((buffer) => ({
          cadesplugin,
          base64Data: toBase64(buffer),
        }))
      )
      .then(({ cadesplugin, base64Data }) =>
        createCadesDetachedSignature(cadesplugin, selectedThumbprint, base64Data)
      )
      .then((signatureBase64) => {
        const signatureBytes = base64ToBytes(signatureBase64)
        const signatureBlob = new Blob([signatureBytes], {
          type: "application/pkcs7-signature",
        })
        const signatureFileName = getSignatureFileName(documentFile.name)

        downloadBlob(signatureBlob, signatureFileName)
        setSignature(`SIG-подпись сформирована и скачана: ${signatureFileName}`)
        setStatus("Файл успешно подписан. SIG-подпись скачана.")
      })
      .catch((error) => {
        const errorMessage = getErrorMessage(error)
        setStatus(`Ошибка создания .sig: ${errorMessage}`)
      })
      .finally(() => setIsSigning(false))
  }

  return (
    <main className="page">
      <header className="hero">
        <p className="hero-kicker">CryptoPro Workspace</p>
        <h1 className="page-title">Центр электронной подписи</h1>
        <p className="hero-subtitle">
          Выберите сертификат, загрузите файл и получите отдельную SIG-подпись.
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
          <span>Результат SIG</span>
        </h2>
        <textarea
          className="signature-output"
          value={signature}
          readOnly
          placeholder="После успешной подписи .sig файл скачается автоматически"
        />
      </section>

      <section className="card">
        <h2 className="card-title">Текущий статус</h2>
        <p className="status">{status}</p>
      </section>

    </main>
  )
}
