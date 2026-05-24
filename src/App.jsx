import { useEffect, useState } from "react";
import logo from "./assets/logo.png";
import { supabase } from "./supabase";
import * as XLSX from "xlsx";

const emptyServiceItem = {
  butiran: "",
  warranty: "",
  harga: "",
};

const emptyForm = {
  nama: "",
  telefon: "",
  peranti: "",
  model: "",
  masalah: "",
  serviceItems: [],
  jenisHarga: "Belum Check",
  anggaranMin: "",
  anggaranMax: "",
  kos: "",
  deposit: "",
  status: "Diterima",
  tarikhAmbil: "",
  masaAmbil: "",
};

function App() {
  const [data, setData] = useState([]);
  const [search, setSearch] = useState("");
  const [form, setForm] = useState(emptyForm);
  const [serviceItem, setServiceItem] = useState(emptyServiceItem);
  const [editId, setEditId] = useState(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [session, setSession] = useState(null);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authLoading, setAuthLoading] = useState(true);
  const [loginLoading, setLoginLoading] = useState(false);

  const rowsPerPage = 5;

  useEffect(() => {
    async function checkSession() {
      const { data } = await supabase.auth.getSession();
      setSession(data.session);
      setAuthLoading(false);
    }

    checkSession();

    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (session) fetchrepairs();
      else setData([]);
    });

    return () => {
      listener.subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (session) fetchrepairs();
  }, [session]);

  async function fetchrepairs() {
    const { data, error } = await supabase
      .from("repairs")
      .select("*")
      .order("id", { ascending: false });

    if (error) {
      console.error("Supabase fetch error:", error);
      alert("Database error masa load data. Check console bro.");
      return;
    }

    setData(data || []);
  }

  async function loginAdmin(e) {
    e.preventDefault();

    if (!email || !password) {
      alert("Masukkan email dan password dulu bro 😆");
      return;
    }

    setLoginLoading(true);

    const { error } = await supabase.auth.signInWithPassword({
      email,
      password,
    });

    setLoginLoading(false);

    if (error) {
      console.error("Login error:", error);
      alert("Login gagal bro. Check email/password.");
      return;
    }

    setEmail("");
    setPassword("");
  }

  async function logoutAdmin() {
    await supabase.auth.signOut();
    setEditId(null);
    setForm(emptyForm);
    setServiceItem(emptyServiceItem);
  }

  function generateRepairCode() {
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let code = "";
    let exists = true;

    while (exists) {
      const randomLetters =
        letters[Math.floor(Math.random() * letters.length)] +
        letters[Math.floor(Math.random() * letters.length)] +
        letters[Math.floor(Math.random() * letters.length)];

      const randomNumbers = String(Math.floor(Math.random() * 1000)).padStart(3, "0");
      code = randomLetters + randomNumbers;
      exists = data.some((item) => item.id === code);
    }

    return code;
  }

  function getServiceTotal(item) {
    const serviceItems = item.serviceItems || [];
    return serviceItems.reduce((total, servis) => total + Number(servis.harga || 0), 0);
  }

  function getJumlahKos(item) {
    const totalService = getServiceTotal(item);
    if (totalService > 0) return totalService;
    return Number(item.kos || 0);
  }

  function kiraBaki(item) {
    const jumlahKos = item.jenisHarga === "Harga Final" ? getJumlahKos(item) : 0;
    const deposit = Number(item.deposit || 0);
    return jumlahKos - deposit;
  }

  function getPaparanHarga(item) {
    if (item.jenisHarga === "Belum Check") return "BELUM DIPERIKSA";

    if (item.jenisHarga === "Anggaran") {
      return `RM ${item.anggaranMin || 0} - RM ${item.anggaranMax || 0}`;
    }

    return `RM ${getJumlahKos(item)}`;
  }

  function getDocumentType(item) {
    if (item.status === "Siap") return "INVOIS";
    if (item.status === "Sudah Ambil") return "RESIT BAYARAN";
    return "RESIT PENERIMAAN PERANTI";
  }

  function getDocumentSub(item) {
    if (item.status === "Siap") return "Invois servis repair untuk makluman jumlah bayaran pelanggan.";
    if (item.status === "Sudah Ambil") return "Bukti bayaran dan pengambilan peranti oleh pelanggan.";
    return "Bukti penerimaan peranti untuk tujuan pemeriksaan / servis.";
  }

  function getTarikhMasaTerima(item) {
    return `${item.tarikh || "-"}
${item.masa || "-"}`;
  }

  function getTarikhAmbil(item) {
    return item.tarikhAmbil || item.tarikh_ambil || item.tarikhambil || "";
  }

  function getMasaAmbil(item) {
    return item.masaAmbil || item.masa_ambil || item.masaambil || "";
  }

  function getTarikhMasaAmbil(item) {
    const tarikhAmbil = getTarikhAmbil(item);
    const masaAmbil = getMasaAmbil(item);
    if (!tarikhAmbil && !masaAmbil) return `-\n-`;
    return `${tarikhAmbil || "-"}
${masaAmbil || "-"}`;
  }


  function exportDatabaseExcel() {
    if (!data || data.length === 0) {
      alert("Tiada data untuk export bro 😆");
      return;
    }

    let totalHarga = 0;
    let totalBayaran = 0;
    let totalBaki = 0;

    const rows = data.map((item) => {
      const serviceItems = item.serviceItems || [];

      const harga =
        item.jenisHarga === "Harga Final"
          ? Number(getJumlahKos(item) || 0)
          : 0;

      const bayaran = Number(item.deposit || 0);

      const baki =
        item.jenisHarga === "Harga Final"
          ? Number(item.baki || 0)
          : 0;

      totalHarga += harga;
      totalBayaran += bayaran;
      totalBaki += baki;

      const butiranServis = serviceItems.length > 0
        ? serviceItems
            .map(
              (servis, index) =>
                `${index + 1}. ${servis.butiran || "-"} | Warranty: ${servis.warranty || "-"} | RM ${servis.harga || 0}`
            )
            .join("\n")
        : "-";

      return {
        "No Repair": item.id || "",
        "Nama": item.nama || "",
        "Telefon": item.telefon || "",
        "Peranti": item.peranti || "",
        "Model": item.model || "",
        "Masalah": item.masalah || "",
        "Tarikh/Masa Terima": getTarikhMasaTerima(item),
        "Tarikh/Masa Ambil": getTarikhMasaAmbil(item),
        "Status": item.status || "",
        "Jenis Harga": item.jenisHarga || "",
        "Harga / Kos": harga,
        "Bayaran / Deposit": bayaran,
        "Baki": baki,
        "Dokumen": getDocumentType(item),
        "Butiran Servis": butiranServis,
      };
    });

    rows.push({});
    rows.push({
      "Jenis Harga": "JUMLAH KESELURUHAN",
      "Harga / Kos": totalHarga,
      "Bayaran / Deposit": totalBayaran,
      "Baki": totalBaki,
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);

    worksheet["!cols"] = [
      { wch: 12 }, // No Repair
      { wch: 22 }, // Nama
      { wch: 16 }, // Telefon
      { wch: 16 }, // Peranti
      { wch: 22 }, // Model
      { wch: 35 }, // Masalah
      { wch: 20 }, // Tarikh/Masa Terima
      { wch: 20 }, // Tarikh/Masa Ambil
      { wch: 25 }, // Status
      { wch: 20 }, // Jenis Harga
      { wch: 18 }, // Harga / Kos
      { wch: 18 }, // Bayaran / Deposit
      { wch: 18 }, // Baki
      { wch: 28 }, // Dokumen
      { wch: 55 }, // Butiran Servis
    ];

    const totalRowNumber = rows.length + 1;
    const totalLabelCell = `J${totalRowNumber}`;
    const totalHargaCell = `K${totalRowNumber}`;
    const totalBayaranCell = `L${totalRowNumber}`;
    const totalBakiCell = `M${totalRowNumber}`;

    [totalLabelCell, totalHargaCell, totalBayaranCell, totalBakiCell].forEach((cell) => {
      if (worksheet[cell]) {
        worksheet[cell].s = {
          font: { bold: true },
        };
      }
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Database Repair");

    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `DQ-Tech-Database-Repair-${today}.xlsx`);
  }

  function tambahServiceItem() {
    if (!serviceItem.butiran) {
      alert("Isi butiran servis dulu bro 😆");
      return;
    }

    const newItems = [...(form.serviceItems || []), serviceItem];
    const newTotal = newItems.reduce((total, item) => total + Number(item.harga || 0), 0);

    setForm({
      ...form,
      serviceItems: newItems,
      kos: newTotal || form.kos,
      jenisHarga: newTotal > 0 ? "Harga Final" : form.jenisHarga,
    });

    setServiceItem(emptyServiceItem);
  }

  function deleteServiceItem(index) {
    const newItems = (form.serviceItems || []).filter((_, i) => i !== index);
    const newTotal = newItems.reduce((total, item) => total + Number(item.harga || 0), 0);

    setForm({
      ...form,
      serviceItems: newItems,
      kos: newTotal || "",
    });
  }

  async function saveRepair() {
    if (!form.nama || !form.telefon || !form.peranti || !form.masalah) {
      alert("Isi nama, telefon, peranti dan masalah dulu bro 😆");
      return;
    }

    const now = new Date();
    const tarikhSekarang = now.toLocaleDateString("ms-MY");
    const masaSekarang = now.toLocaleTimeString("en-MY", {
      hour: "2-digit",
      minute: "2-digit",
    });
    const existingRepair = editId ? data.find((item) => item.id === editId) : null;

const finalForm = {
  ...form,
  kos: Number(getServiceTotal(form) > 0 ? getServiceTotal(form) : form.kos || 0),
  deposit: Number(form.deposit || 0),
  tarikhAmbil:
    form.status === "Sudah Ambil"
      ? form.tarikhAmbil || existingRepair?.tarikhAmbil || tarikhSekarang
      : form.tarikhAmbil || existingRepair?.tarikhAmbil || "",
  masaAmbil:
    form.status === "Sudah Ambil"
      ? form.masaAmbil || existingRepair?.masaAmbil || masaSekarang
      : form.masaAmbil || existingRepair?.masaAmbil || "",
};

    if (editId) {
      const { error } = await supabase
        .from("repairs")
        .update({
          ...finalForm,
          kos: Number(finalForm.kos || 0),
          deposit: Number(finalForm.deposit || 0),
          baki: kiraBaki(finalForm),
          anggaranMin: finalForm.anggaranMin || "",
          anggaranMax: finalForm.anggaranMax || "",
          serviceItems: finalForm.serviceItems || [],
          tarikhAmbil: finalForm.tarikhAmbil || "",
          masaAmbil: finalForm.masaAmbil || "",
        })
        .eq("id", editId);

      if (error) {
        console.error("Supabase update error:", error);
        alert("Update gagal bro. Check table column Supabase.");
        return;
      }

      await fetchrepairs();
      setForm(emptyForm);
      setServiceItem(emptyServiceItem);
      setEditId(null);
      alert("Repair berjaya dikemaskini 🔥");
      return;
    }

    const newRepair = {
      id: generateRepairCode(),
      tarikh: tarikhSekarang,
      masa: masaSekarang,
      ...finalForm,
      baki: kiraBaki(finalForm),
    };

    const { error } = await supabase.from("repairs").insert([
      {
        ...newRepair,
        kos: Number(newRepair.kos || 0),
        deposit: Number(newRepair.deposit || 0),
        baki: kiraBaki(finalForm),
        anggaranMin: newRepair.anggaranMin || "",
        anggaranMax: newRepair.anggaranMax || "",
        serviceItems: newRepair.serviceItems || [],
        tarikhAmbil: newRepair.tarikhAmbil || "",
        masaAmbil: newRepair.masaAmbil || "",
      },
    ]);

    if (error) {
      console.error("Supabase insert error:", error);
      alert("Simpan gagal bro. Check table column Supabase.");
      return;
    }

    await fetchrepairs();
    setForm(emptyForm);
    setServiceItem(emptyServiceItem);
    setCurrentPage(1);
    alert("Repair berjaya disimpan 🔥");
  }

  function editRepair(item) {
    setEditId(item.id);

    setForm({
      nama: item.nama || "",
      telefon: item.telefon || "",
      peranti: item.peranti || "",
      model: item.model || "",
      masalah: item.masalah || "",
      serviceItems: item.serviceItems || [],
      jenisHarga: item.jenisHarga || "Belum Check",
      anggaranMin: item.anggaranMin || "",
      anggaranMax: item.anggaranMax || "",
      kos: item.kos || "",
      deposit: item.deposit || "",
      status: item.status || "Diterima",
      tarikhAmbil: getTarikhAmbil(item),
      masaAmbil: getMasaAmbil(item),
    });

    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function deleteRepair(id) {
    const confirmDelete = window.confirm(`Confirm padam repair ${id}?`);

    if (!confirmDelete) return;

    const { error } = await supabase
      .from("repairs")
      .delete()
      .eq("id", id);

    if (error) {
      console.error("Supabase delete error:", error);
      alert("Padam gagal bro. Check Supabase policy/table.");
      return;
    }

    if (editId === id) {
      setEditId(null);
      setForm(emptyForm);
      setServiceItem(emptyServiceItem);
    }

    await fetchrepairs();
    alert("Repair berjaya dipadam 🗑️");
  }

  function cancelEdit() {
    setEditId(null);
    setForm(emptyForm);
    setServiceItem(emptyServiceItem);
  }

  function handleSearch(value) {
    setSearch(value);
    setCurrentPage(1);
  }

  const result = data.filter((item) =>
    JSON.stringify(item).toLowerCase().includes(search.toLowerCase())
  );

  const totalPages = Math.max(1, Math.ceil(result.length / rowsPerPage));
  const startIndex = (currentPage - 1) * rowsPerPage;
  const paginatedResult = result.slice(startIndex, startIndex + rowsPerPage);

  const latest = editId ? data.find((item) => item.id === editId) || data[0] : data[0];

  if (authLoading) {
    return (
      <div style={styles.loginPage}>
        <div style={styles.loginCard}>Loading system...</div>
      </div>
    );
  }

  if (!session) {
    return (
      <div style={styles.loginPage}>
        <form style={styles.loginCard} onSubmit={loginAdmin}>
          <img src={logo} alt="DQ Tech" style={styles.loginLogo} />
          <h1 style={styles.loginTitle}>DQ Tech Repair System</h1>
          <p style={styles.loginSub}>Admin Login</p>

          <label style={styles.label}>Email</label>
          <input
            style={styles.input}
            type="email"
            placeholder="admin email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />

          <label style={styles.label}>Password</label>
          <input
            style={styles.input}
            type="password"
            placeholder="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />

          <button style={styles.loginBtn} type="submit" disabled={loginLoading}>
            {loginLoading ? "LOGIN..." : "LOGIN ADMIN"}
          </button>
        </form>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      <div className="shop-header" style={styles.header}>
        <div>
          <h1 style={styles.headerTitle}>DQ Tech Repair System</h1>
          <p style={styles.headerSub}>Invois • Payment Resit • Servis Repair Form</p>
        </div>

        <div style={styles.headerActions}>
          <div style={styles.nextNo}>
            <div>{editId ? "Sedang Edit" : "No. Repair"}</div>
            <b>{editId || "Auto Generate"}</b>
          </div>

          <button style={styles.logoutBtn} onClick={logoutAdmin}>
            LOGOUT
          </button>
        </div>
      </div>

      <div style={styles.mainGrid}>
        <div style={styles.card}>
          <h2 style={styles.formTitle}>
            {editId ? "✏️ EDIT REPAIR FORM" : "🛠️ SERVIS REPAIR FORM"}
          </h2>

          <div style={styles.titleLine}></div>

          {editId && (
            <div style={styles.editNotice}>
              Sedang edit repair: <b>{editId}</b>
            </div>
          )}

          <Label text="Nama Pelanggan" />
          <input style={styles.input} placeholder="Nama pelanggan" value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} />

          <Label text="No. Telefon" />
          <input style={styles.input} placeholder="No telefon" value={form.telefon} onChange={(e) => setForm({ ...form, telefon: e.target.value })} />

          <Label text="Jenis Peranti" />
          <input style={styles.input} placeholder="Contoh: Laptop / Desktop / PS5" value={form.peranti} onChange={(e) => setForm({ ...form, peranti: e.target.value })} />

          <Label text="Model / Jenama" />
          <input style={styles.input} placeholder="Contoh: Dell Optiplex / Sony PS5" value={form.model} onChange={(e) => setForm({ ...form, model: e.target.value })} />

          <Label text="Masalah Peranti" />
          <textarea style={styles.textarea} placeholder="Contoh: Laptop tidak hidup / no display / liquid damage" value={form.masalah} onChange={(e) => setForm({ ...form, masalah: e.target.value })} />

          <div style={styles.serviceBox}>
            <h3 style={styles.serviceTitle}>BUTIRAN SERVIS / ITEM REPAIR</h3>

            <Label text="Butiran Servis" />
            <input
              style={styles.input}
              placeholder='Contoh: SSD 2.5" 256GB - S/N BTPY708605DC256D'
              value={serviceItem.butiran}
              onChange={(e) => setServiceItem({ ...serviceItem, butiran: e.target.value })}
            />

            <div style={styles.twoInput}>
              <div>
                <Label text="Warranty" />
                <input
                  style={styles.input}
                  placeholder="Contoh: 6 Bulan"
                  value={serviceItem.warranty}
                  onChange={(e) => setServiceItem({ ...serviceItem, warranty: e.target.value })}
                />
              </div>

              <div>
                <Label text="Harga (RM)" />
                <input
                  style={styles.input}
                  placeholder="Contoh: 280"
                  value={serviceItem.harga}
                  onChange={(e) => setServiceItem({ ...serviceItem, harga: e.target.value })}
                />
              </div>
            </div>

            <button style={styles.addItemBtn} onClick={tambahServiceItem}>
              + TAMBAH ITEM SERVIS
            </button>

            {(form.serviceItems || []).length > 0 && (
              <table style={styles.itemTable}>
                <thead>
                  <tr>
                    <th>No</th>
                    <th>Butiran</th>
                    <th>Warranty</th>
                    <th>Harga</th>
                    <th>Action</th>
                  </tr>
                </thead>

                <tbody>
                  {form.serviceItems.map((item, index) => (
                    <tr key={index}>
                      <td>{index + 1}</td>
                      <td style={{ textAlign: "left" }}>{item.butiran}</td>
                      <td>{item.warranty || "-"}</td>
                      <td>RM {item.harga || 0}</td>
                      <td>
                        <button style={styles.deleteBtn} onClick={() => deleteServiceItem(index)}>
                          X
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          <Label text="Jenis Harga" />
          <select style={styles.input} value={form.jenisHarga} onChange={(e) => setForm({ ...form, jenisHarga: e.target.value })}>
            <option>Belum Check</option>
            <option>Anggaran</option>
            <option>Harga Final</option>
          </select>

          {form.jenisHarga === "Anggaran" && (
            <div style={styles.twoInput}>
              <div>
                <Label text="Anggaran Min (RM)" />
                <input style={styles.input} placeholder="Contoh: 200" value={form.anggaranMin} onChange={(e) => setForm({ ...form, anggaranMin: e.target.value })} />
              </div>

              <div>
                <Label text="Anggaran Max (RM)" />
                <input style={styles.input} placeholder="Contoh: 400" value={form.anggaranMax} onChange={(e) => setForm({ ...form, anggaranMax: e.target.value })} />
              </div>
            </div>
          )}

          {form.jenisHarga === "Harga Final" && (
            <>
              <Label text="Harga Final / Jumlah Kos (RM)" />
              <input style={styles.input} placeholder="Contoh: 1000" value={form.kos} onChange={(e) => setForm({ ...form, kos: e.target.value })} />
            </>
          )}

          <Label text="Deposit / Bayaran (RM)" />
          <input style={styles.input} placeholder="Contoh: 100" value={form.deposit} onChange={(e) => setForm({ ...form, deposit: e.target.value })} />

          <Label text="Status Repair" />
          <select style={styles.input} value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
            <option>Diterima</option>
            <option>Dalam Pemeriksaan</option>
            <option>Menunggu Confirmation Customer</option>
            <option>Menunggu Parts</option>
            <option>Dalam Repair</option>
            <option>Siap</option>
            <option>Sudah Ambil</option>
            <option>Cancel Repair</option>
          </select>

          <button style={editId ? styles.updateBtn : styles.saveBtn} onClick={saveRepair}>
            {editId ? "✅ UPDATE REPAIR" : "💾 SIMPAN REPAIR"}
          </button>

          {editId && (
            <button style={styles.cancelBtn} onClick={cancelEdit}>
              ❌ CANCEL EDIT
            </button>
          )}
        </div>

        <div style={styles.card}>
          {latest ? (
            <div className="receipt-print" style={styles.receipt}>
              <div className="shopHeader" style={styles.shopHeader}>
                <div className="logo-container" style={styles.logoContainer}>
                  <img src={logo} alt="DQ Tech" style={styles.logo} />
                </div>

                <div style={styles.shopInfo}>

  <div style={styles.shopTitleRow}>
    <h1 style={styles.shopName}>
      DQ TECH RESOURCES
    </h1>

    <span style={styles.regNo}>
      (NS0274287-U)
    </span>
  </div>

  <p style={styles.shopText}>
    NO 12 JALAN TAS 30
  </p>

  <p style={styles.shopText}>
    TAMAN TAS, 25150 KUANTAN
  </p>

  <p style={styles.shopText}>
    TEL : 017-5155323
  </p>

  <p style={styles.shopText}>
    www.facebook.com/dqtechktn
  </p>

</div>
              </div>

              <div style={styles.blueLine}></div>

              <div className="receiptTitleRow" style={styles.receiptTitleRow}>
                <div>
                  <h2 style={styles.receiptTitle}>{getDocumentType(latest)}</h2>
                  <p style={styles.receiptSub}>{getDocumentSub(latest)}</p>
                </div>

                <div style={styles.receiptNo}>{latest.id}</div>
              </div>

              <div className="infoGrid" style={styles.infoGrid}>
                <div style={styles.infoBox}>
                  <h3 style={styles.boxTitle}>MAKLUMAT PELANGGAN</h3>
                  <Info label="Nama Pelanggan" value={latest.nama} />
                  <Info label="No. Telefon" value={latest.telefon} />
                  <Info label="Tarikh Terima" value={latest.tarikh} />
                  <Info label="Masa Terima" value={latest.masa || "-"} />
                </div>

                <div style={styles.infoBox}>
                  <h3 style={styles.boxTitle}>MAKLUMAT PERANTI</h3>
                  <Info label="Jenis Peranti" value={latest.peranti} />
                  <Info label="Model / Jenama" value={latest.model || "-"} />
                  <Info label="Status Repair" value={latest.status} blue />
                </div>
              </div>

              <div className="problemBox" style={styles.problemBox}>
                <h3 style={styles.boxTitle}>MASALAH PERANTI</h3>
                <p>{latest.masalah}</p>
              </div>

              {(latest.serviceItems || []).length > 0 && (
                <div className="serviceReceiptBox" style={styles.serviceReceiptBox}>
                  <h3 style={styles.boxTitle}>BUTIRAN SERVIS</h3>

                  <table style={styles.receiptItemTable}>
                    <thead>
                      <tr>
                        <th>No</th>
                        <th>Butiran Servis</th>
                        <th>Warranty</th>
                        <th>Harga</th>
                      </tr>
                    </thead>

                    <tbody>
                      {latest.serviceItems.map((item, index) => (
                        <tr key={index}>
                          <td>{index + 1}</td>
                          <td style={{ textAlign: "left" }}>{item.butiran}</td>
                          <td>{item.warranty || "-"}</td>
                          <td>RM {item.harga || 0}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              <div className="paymentBox" style={styles.paymentBox}>
                <h3 style={styles.boxTitle}>
                  {latest.status === "Sudah Ambil" ? "RINGKASAN BAYARAN" : "RINGKASAN KOS"}
                </h3>

                <PayRow label="Status Harga" value={latest.jenisHarga} />
                <PayRow label={latest.jenisHarga === "Harga Final" ? "Jumlah Kos" : "Anggaran Kos"} value={getPaparanHarga(latest)} />

                <PayRow
                  label={latest.status === "Sudah Ambil" ? "Jumlah Dibayar" : "Deposit / Bayaran"}
                  value={`RM ${latest.deposit || 0}`}
                />

                <div style={latest.status === "Sudah Ambil" && Number(latest.baki || 0) <= 0 ? styles.paidRow : styles.balanceRow}>
                  <span>
                    {latest.status === "Sudah Ambil"
                      ? Number(latest.baki || 0) <= 0
                        ? "STATUS BAYARAN"
                        : "BAKI"
                      : latest.status === "Siap"
                      ? "BAKI PERLU DIBAYAR"
                      : "BAKI"}
                  </span>

                  <b>
                    {latest.status === "Sudah Ambil" && Number(latest.baki || 0) <= 0
                      ? "PAID / SELESAI"
                      : latest.jenisHarga === "Harga Final"
                      ? `RM ${latest.baki}`
                      : "BELUM DITENTUKAN"}
                  </b>
                </div>
              </div>

              <div className="noteBox" style={styles.noteBox}>
                <b>Nota:</b>{" "}
                {latest.status === "Siap"
                  ? "Peranti telah siap. Sila jelaskan baki bayaran sebelum atau semasa pengambilan peranti."
                  : latest.status === "Sudah Ambil"
                  ? "Peranti telah diserahkan kepada pelanggan dan bayaran telah direkodkan berdasarkan maklumat di atas."
                  : "Harga yang dipaparkan adalah berdasarkan status semasa. Jika peranti belum diperiksa sepenuhnya, kos sebenar mungkin berubah selepas pemeriksaan lanjut dan akan dimaklumkan kepada pelanggan sebelum kerja repair diteruskan."}
              </div>

              <div style={styles.thanks}>Terima kasih kerana menggunakan servis DQ Tech 😄</div>

              <div style={styles.iconFooter}>
                <span>🛡️ Servis Berkualiti</span>
                <span>|</span>
                <span>⚙️ Pakar & Berpengalaman</span>
                <span>|</span>
                <span>👍 Kepuasan Pelanggan Diutamakan</span>
              </div>

              <button style={styles.printBtn} onClick={() => window.print()}>
                🖨️ CETAK / SAVE PDF
              </button>
            </div>
          ) : (
            <div style={styles.emptyReceipt}>Belum ada data. Simpan repair dulu.</div>
          )}
        </div>
      </div>

      <div style={styles.databaseCard}>
        <div style={styles.databaseHeader}>
          <h2 style={styles.databaseTitle}>SENARAI REPAIR / DATABASE</h2>

          <input
            style={styles.searchInput}
            placeholder="Cari nama / telefon / model / no repair..."
            value={search}
            onChange={(e) => handleSearch(e.target.value)}
          />
        </div>

        <table style={styles.table}>
          <thead>
            <tr>
              <th>No Repair</th>
              <th>Nama</th>
              <th>Telefon</th>
              <th>Peranti</th>
              <th>Model</th>
              <th>Tarikh/Masa Terima</th>
              <th>Tarikh/Masa Ambil</th>
              <th>Status</th>
              <th>Harga</th>
              <th>Bayaran</th>
              <th>Baki</th>
              <th>Dokumen</th>
              <th>Action</th>
            </tr>
          </thead>

          <tbody>
            {paginatedResult.length > 0 ? (
              paginatedResult.map((item) => (
                <tr key={item.id}>
                  <td><b>{item.id}</b></td>
                  <td>{item.nama}</td>
                  <td>{item.telefon}</td>
                  <td>{item.peranti}</td>
                  <td>{item.model}</td>
                  <td>
                    <div>{item.tarikh || "-"}</div>
                    <div>{item.masa || "-"}</div>
                  </td>
                  <td>
                    <div>{getTarikhAmbil(item) || "-"}</div>
                    <div>{getMasaAmbil(item) || "-"}</div>
                  </td>
                  <td><span style={styles.statusPill}>{item.status}</span></td>
                  <td>{getPaparanHarga(item)}</td>
                  <td>RM {item.deposit || 0}</td>
                  <td style={item.jenisHarga === "Harga Final" && item.baki > 0 ? styles.redText : styles.greenText}>
                    {item.jenisHarga === "Harga Final" ? `RM ${item.baki}` : "Belum Ditentukan"}
                  </td>
                  <td><b>{getDocumentType(item)}</b></td>
                  <td>
                    <div style={styles.actionBox}>
                      <button style={styles.editBtn} onClick={() => editRepair(item)}>
                        EDIT
                      </button>

                      <button style={styles.deleteRepairBtn} onClick={() => deleteRepair(item.id)}>
                        PADAM
                      </button>
                    </div>
                  </td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="14" style={styles.noData}>
                  Tiada data dijumpai.
                </td>
              </tr>
            )}
          </tbody>
        </table>

        <div style={styles.paginationBox}>
          <div style={styles.paginationInfo}>
            Paparan {result.length === 0 ? 0 : startIndex + 1} - {Math.min(startIndex + rowsPerPage, result.length)} daripada {result.length} repair
          </div>

          <div style={styles.paginationButtons}>
            <button style={styles.excelBtn} onClick={exportDatabaseExcel}>
              EXPORT EXCEL
            </button>

            <button
              style={currentPage === 1 ? styles.pageBtnDisabled : styles.pageBtn}
              disabled={currentPage === 1}
              onClick={() => setCurrentPage(1)}
            >
              First
            </button>

            <button
              style={currentPage === 1 ? styles.pageBtnDisabled : styles.pageBtn}
              disabled={currentPage === 1}
              onClick={() => setCurrentPage(currentPage - 1)}
            >
              Previous
            </button>

            <span style={styles.pageNumber}>
              Page {currentPage} / {totalPages}
            </span>

            <button
              style={currentPage === totalPages ? styles.pageBtnDisabled : styles.pageBtn}
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage(currentPage + 1)}
            >
              Next
            </button>

            <button
              style={currentPage === totalPages ? styles.pageBtnDisabled : styles.pageBtn}
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage(totalPages)}
            >
              Last
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function Label({ text }) {
  return <label style={styles.label}>{text}</label>;
}

function Info({ label, value, blue }) {
  return (
    <div style={styles.infoRow}>
      <span>{label}</span>
      <b style={blue ? styles.blueText : {}}>{value}</b>
    </div>
  );
}

function PayRow({ label, value }) {
  return (
    <div style={styles.payRow}>
      <span>{label}</span>
      <b>{value}</b>
    </div>
  );
}

const styles = {
  loginPage: { minHeight: "100vh", background: "linear-gradient(135deg, #020617, #0f172a)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "Arial, sans-serif" },
  loginCard: { width: 380, background: "white", borderRadius: 10, padding: 28, boxShadow: "0 12px 35px #0005", textAlign: "center", color: "#0f172a" },
  loginLogo: { width: 95, height: 95, objectFit: "contain", background: "#000", borderRadius: 8, marginBottom: 15 },
  loginTitle: { margin: 0, fontSize: 24, color: "#0057c2" },
  loginSub: { margin: "8px 0 20px", color: "#64748b", fontWeight: "bold" },
  loginBtn: { marginTop: 20, width: "100%", padding: 14, border: 0, borderRadius: 5, background: "#0057c2", color: "white", fontWeight: "bold", fontSize: 15, cursor: "pointer" },
  page: { padding: 18, background: "#f8fafc", fontFamily: "Arial, sans-serif", minHeight: "100vh", color: "#0f172a" },
  header: { background: "linear-gradient(135deg, #020617, #0f172a)", color: "white", borderRadius: 8, padding: "22px 30px", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 25, boxShadow: "0 6px 18px #0002" },
  headerTitle: { margin: 0, fontSize: 28 },
  headerSub: { margin: "6px 0 0", fontWeight: "bold" },
  headerActions: { display: "flex", alignItems: "center", gap: 12 },
  nextNo: { background: "linear-gradient(135deg, #0057c2, #006ee6)", padding: "14px 45px", borderRadius: 8, textAlign: "center", fontSize: 14 },
  logoutBtn: { background: "#dc2626", color: "white", border: 0, borderRadius: 6, padding: "12px 16px", fontWeight: "bold", cursor: "pointer" },
  mainGrid: { display: "grid", gridTemplateColumns: "430px 1fr", gap: 25, alignItems: "start" },
  card: { background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 25, boxShadow: "0 4px 16px #0001" },
  formTitle: { color: "#0057c2", margin: 0, fontSize: 22 },
  titleLine: { width: 180, height: 3, background: "#006ee6", margin: "14px 0 20px" },
  editNotice: { background: "#eff6ff", color: "#0057c2", border: "1px solid #bfdbfe", borderRadius: 6, padding: 10, fontWeight: "bold", marginBottom: 15 },
  label: { display: "block", fontWeight: "bold", fontSize: 14, marginBottom: 8, marginTop: 14 },
  input: { width: "100%", padding: "12px", borderRadius: 5, border: "1px solid #cbd5e1", fontSize: 14, boxSizing: "border-box" },
  textarea: { width: "100%", height: 115, padding: "12px", borderRadius: 5, border: "1px solid #cbd5e1", fontSize: 14, boxSizing: "border-box" },
  twoInput: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 18 },
  serviceBox: { background: "#f8fafc", border: "1px solid #dbeafe", borderRadius: 8, padding: 14, marginTop: 18 },
  serviceTitle: { margin: 0, color: "#0057c2", fontSize: 15, textAlign: "center" },
  addItemBtn: { marginTop: 15, width: "100%", padding: 12, border: 0, borderRadius: 5, background: "#0f172a", color: "white", fontWeight: "bold", cursor: "pointer" },
  deleteBtn: { background: "#dc2626", color: "white", border: 0, borderRadius: 4, padding: "5px 8px", fontWeight: "bold", cursor: "pointer" },
  itemTable: { width: "100%", borderCollapse: "collapse", marginTop: 15, fontSize: 12, background: "white" },
  saveBtn: { marginTop: 20, width: "100%", padding: 14, border: 0, borderRadius: 5, background: "#16a34a", color: "white", fontWeight: "bold", fontSize: 15, cursor: "pointer" },
  updateBtn: { marginTop: 20, width: "100%", padding: 14, border: 0, borderRadius: 5, background: "#0057c2", color: "white", fontWeight: "bold", fontSize: 15, cursor: "pointer" },
  cancelBtn: { marginTop: 10, width: "100%", padding: 13, border: 0, borderRadius: 5, background: "#dc2626", color: "white", fontWeight: "bold", fontSize: 14, cursor: "pointer" },
  receipt: { background: "white", color: "#0f172a", width: "100%", zoom: "78%" },
shopHeader: {
  display: "flex",
  alignItems: "flex-start",
  justifyContent: "flex-start",
  width: "100%",
  gap: 15,
  padding: "5px 0"
},

logoContainer: {
  flex: "0 0 230px",
  display: "flex",
  justifyContent: "center",
  alignItems: "stretch"
},

logo: {
  width: 220,
  height: 180,
  objectFit: "cover",
  background: "transparent"
},

shopInfo: {
  flex: 1,
  textAlign: "left",
  paddingLeft: 8
},

shopTitleRow: {
  display: "flex",
  alignItems: "center",
  gap: 16,
  marginBottom: 18,
},

shopName: {
  margin: 0,
  color: "#000",
  fontSize: 28,
  fontWeight: "700",
  letterSpacing: 0.5
},

regNo: {
  fontWeight: "700",
  fontSize: 14,
  color: "#0f172a"
},

shopText: {
  margin: "10px 0",
  fontWeight: "600",
  fontSize: 14
},

blueLine: {
  height: 4,
  background: "#006ee6",
  margin: "10px 0 16px"
},
  receiptTitleRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 },
  receiptTitle: { margin: 0, color: "#0057c2", fontSize: 24 },
  receiptSub: { margin: "6px 0 0", color: "#64748b", fontSize: 12 },
  receiptNo: { background: "#0057c2", color: "white", padding: "10px 18px", borderRadius: 6, fontWeight: "bold", fontSize: 18 },
  infoGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 },
  infoBox: { border: "1px solid #cbd5e1", borderRadius: 8, padding: 12 },
  boxTitle: { textAlign: "center", color: "#0057c2", margin: "0 0 12px", fontSize: 15 },
  infoRow: { display: "flex", justifyContent: "space-between", marginBottom: 10, fontSize: 13, gap: 10 },
  blueText: { color: "#0057c2" },
  problemBox: { border: "1px solid #cbd5e1", borderRadius: 8, padding: 12, marginTop: 14, textAlign: "center" },
  serviceReceiptBox: { border: "1px solid #cbd5e1", borderRadius: 8, padding: 12, marginTop: 14 },
  receiptItemTable: { width: "100%", borderCollapse: "collapse", textAlign: "center", fontSize: 13 },
  paymentBox: { border: "1px solid #cbd5e1", borderRadius: 8, marginTop: 14, overflow: "hidden" },
  payRow: { display: "flex", justifyContent: "space-between", padding: "10px 18px", fontSize: 13, gap: 10 },
  balanceRow: { display: "flex", justifyContent: "space-between", padding: "12px 18px", background: "#eff6ff", color: "#0057c2", fontWeight: "bold", fontSize: 18, gap: 10 },
  paidRow: { display: "flex", justifyContent: "space-between", padding: "12px 18px", background: "#dcfce7", color: "#15803d", fontWeight: "bold", fontSize: 18, gap: 10 },
  noteBox: { marginTop: 14, border: "1px solid #fde68a", background: "#fffbeb", color: "#78350f", borderRadius: 8, padding: 10, fontSize: 12, lineHeight: 1.35 },
  thanks: { textAlign: "center", margin: "10px 0", fontSize: 13 },
  iconFooter: { display: "flex", justifyContent: "center", gap: 14, color: "#475569", fontSize: 13, marginTop: 8, marginBottom: 8 },
  printBtn: { display: "block", margin: "0 auto", background: "#0057c2", color: "white", border: 0, padding: "14px 35px", borderRadius: 5, fontWeight: "bold", fontSize: 15, cursor: "pointer" },
  emptyReceipt: { textAlign: "center", padding: 80, color: "#64748b" },
  databaseCard: { marginTop: 25, background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 22, boxShadow: "0 4px 16px #0001", overflowX: "auto" },
  databaseHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 20, marginBottom: 15 },
  databaseTitle: { color: "#0057c2", fontSize: 18, margin: 0 },
  searchInput: { width: 520, padding: 12, borderRadius: 5, border: "1px solid #cbd5e1" },
  table: { width: "100%", borderCollapse: "collapse", textAlign: "center", fontSize: 13 },
  statusPill: { background: "#dcfce7", color: "#15803d", padding: "7px 12px", borderRadius: 7, fontWeight: "bold" },
  actionBox: { display: "flex", gap: 6, justifyContent: "center", alignItems: "center" },
  editBtn: { background: "#0057c2", color: "white", border: 0, borderRadius: 5, padding: "8px 12px", fontWeight: "bold", cursor: "pointer" },
  deleteRepairBtn: { background: "#dc2626", color: "white", border: 0, borderRadius: 5, padding: "8px 12px", fontWeight: "bold", cursor: "pointer" },
  redText: { color: "#dc2626", fontWeight: "bold" },
  greenText: { color: "#16a34a", fontWeight: "bold" },
  paginationBox: { marginTop: 18, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 15, flexWrap: "wrap" },
  paginationInfo: { fontSize: 13, color: "#475569", fontWeight: "bold" },
  paginationButtons: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" },
  excelBtn: { background: "#16a34a", color: "white", border: 0, borderRadius: 5, padding: "9px 13px", fontWeight: "bold", cursor: "pointer" },
  pageBtn: { background: "#0057c2", color: "white", border: 0, borderRadius: 5, padding: "9px 13px", fontWeight: "bold", cursor: "pointer" },
  pageBtnDisabled: { background: "#cbd5e1", color: "#64748b", border: 0, borderRadius: 5, padding: "9px 13px", fontWeight: "bold", cursor: "not-allowed" },
  pageNumber: { fontWeight: "bold", color: "#0f172a", padding: "0 8px" },
  noData: { padding: 25, color: "#64748b", fontWeight: "bold" },
};

const printStyle = document.createElement("style");
printStyle.innerHTML = `
table th,
table td {
  border: 1px solid #cbd5e1;
  padding: 9px;
}

@media print {
  @page {
    size: A4;
    margin: 3mm;
  }

  body * {
    visibility: hidden;
  }

  .receipt-print,
  .receipt-print * {
    visibility: visible;
  }

  .receipt-print {
    position: absolute;
    left: 50%;
    top: 0;
    width: 190mm !important;
    max-height: 291mm !important;
    transform: translateX(-50%) scale(0.96);
    transform-origin: top center;
    background: white !important;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

/* FIX LOGO PRINT */
.shopHeader {
  align-items: flex-start !important;
}

.logoContainer {
  height: 135px !important;
  overflow: hidden !important;
}

.logoContainer img {
  width: 180px !important;
  height: 135px !important;
  object-fit: contain !important;
}

  .problemBox,
  .serviceReceiptBox,
  .paymentBox,
  .noteBox,
  .infoGrid,
  .receiptTitleRow {
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }

  .paymentBox {
    display: block !important;
  }

  table,
  tr,
  td,
  th {
    break-inside: avoid !important;
    page-break-inside: avoid !important;
  }

  button {
    display: none !important;
  }
}
`;

document.head.appendChild(printStyle);

export default App;