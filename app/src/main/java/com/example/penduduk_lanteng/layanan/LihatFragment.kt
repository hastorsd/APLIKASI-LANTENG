package com.example.penduduk_lanteng.layanan

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Toast
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import androidx.fragment.app.Fragment
import androidx.lifecycle.lifecycleScope
import androidx.navigation.fragment.findNavController
import com.example.penduduk_lanteng.DB.AppDatabase
import com.example.penduduk_lanteng.DB.dao.PendudukDao
import com.example.penduduk_lanteng.DB.entity.Penduduk
import com.example.penduduk_lanteng.R
import com.example.penduduk_lanteng.databinding.FragmentLihatBinding
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream
import java.io.InputStream
import java.io.OutputStream

class LihatFragment : Fragment() {

    private var _binding: FragmentLihatBinding? = null
    private val binding get() = _binding!!
    private val CREATE_FILE_REQUEST_CODE = 1
    private val IMPORT_FILE_REQUEST_CODE = 2
    private lateinit var pendudukDao: PendudukDao

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View? {
        _binding = FragmentLihatBinding.inflate(inflater, container, false)

        // Inisialisasi PendudukDao menggunakan AppDatabase
        val db = AppDatabase.getInstance(requireContext())
        pendudukDao = db?.pendudukDao() ?: throw IllegalStateException("Database not initialized")

        binding.rt1.setOnClickListener {
            findNavController().navigate(R.id.action_lihatFragment_to_data1Fragment)
        }

        binding.rt2.setOnClickListener {
            findNavController().navigate(R.id.action_lihatFragment_to_data2Fragment)
        }

        binding.rt3.setOnClickListener {
            findNavController().navigate(R.id.action_lihatFragment_to_data3Fragment)
        }

        binding.rt4.setOnClickListener {
            findNavController().navigate(R.id.action_lihatFragment_to_data4Fragment)
        }

        binding.buttonExportToExcel.setOnClickListener {
            exportToDownloads()
        }

        binding.buttonImportFromExcel.setOnClickListener {
            openFilePicker()
        }

        return binding.root
    }

    private fun requestPermissions() {
        if (ContextCompat.checkSelfPermission(
                requireContext(),
                Manifest.permission.WRITE_EXTERNAL_STORAGE
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(
                requireActivity(),
                arrayOf(Manifest.permission.WRITE_EXTERNAL_STORAGE),
                1
            )
        }
    }

    private fun exportToDownloads() {
        lifecycleScope.launch {
            val downloadsDir =
                Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
            val file = File(downloadsDir, "DataPenduduk.xls")
            val uri = Uri.fromFile(file)

            writeDataToExcel(uri)
        }
    }

    private suspend fun writeDataToExcel(uri: Uri) {
        try {
            val outputStream: OutputStream = FileOutputStream(File(uri.path!!))
            val workbook = jxl.Workbook.createWorkbook(outputStream)
            val sheet = workbook.createSheet("Data Penduduk", 0)

            // Membuat Header
            val headers = listOf(
                "NIK",
                "Nama",
                "Alias",
                "Tempat Lahir",
                "Tanggal Lahir",
                "Agama",
                "Pekerjaan",
                "Kelamin",
                "RT",
                "Status",
                "Hidup"
            )
            headers.forEachIndexed { index, header ->
                sheet.addCell(jxl.write.Label(index, 0, header))
            }

            // Ambil data penduduk dari database
            val pendudukList = pendudukDao.getAllPenduduk()
            pendudukList.forEachIndexed { rowIndex, penduduk ->
                sheet.addCell(jxl.write.Label(0, rowIndex + 1, penduduk.nik))
                sheet.addCell(jxl.write.Label(1, rowIndex + 1, penduduk.nama))
                sheet.addCell(jxl.write.Label(2, rowIndex + 1, penduduk.alias))
                sheet.addCell(jxl.write.Label(3, rowIndex + 1, penduduk.tempat_lahir))
                sheet.addCell(jxl.write.Label(4, rowIndex + 1, penduduk.tanggal_lahir))
                sheet.addCell(jxl.write.Label(5, rowIndex + 1, penduduk.agama))
                sheet.addCell(jxl.write.Label(6, rowIndex + 1, penduduk.pekerjaan))
                sheet.addCell(jxl.write.Label(7, rowIndex + 1, penduduk.kelamin))
                sheet.addCell(jxl.write.Label(8, rowIndex + 1, penduduk.rt))
                sheet.addCell(jxl.write.Label(9, rowIndex + 1, penduduk.status))
                sheet.addCell(jxl.write.Label(10, rowIndex + 1, penduduk.hidup))
            }

            // Menyimpan Workbook
            workbook.write()
            workbook.close()
            outputStream.close()

            Toast.makeText(requireContext(), "Data berhasil diekspor", Toast.LENGTH_LONG).show()
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(requireContext(), "Gagal mengekspor data", Toast.LENGTH_SHORT).show()
        }
    }

    private fun openFilePicker() {
        val intent = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
            addCategory(Intent.CATEGORY_OPENABLE)
            type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" // Untuk .xlsx
            putExtra(
                Intent.EXTRA_MIME_TYPES, arrayOf(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
                    "application/vnd.ms-excel" // .xls
                )
            )
        }
        startActivityForResult(intent, IMPORT_FILE_REQUEST_CODE)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)

        if (requestCode == IMPORT_FILE_REQUEST_CODE && resultCode == android.app.Activity.RESULT_OK && data != null) {
            val uri = data.data
            uri?.let { readExcelFile(it) }
        }
    }

    private fun readExcelFile(uri: Uri) {
        try {
            val inputStream: InputStream? = requireContext().contentResolver.openInputStream(uri)
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0)

            // Iterasi melalui setiap baris di sheet
            for (row in sheet) {
                val nik = row.getCell(0).stringCellValue
                val nama = row.getCell(1).stringCellValue
                val alias = row.getCell(2).stringCellValue
                val tempatLahir = row.getCell(3).stringCellValue
                val tanggalLahir = row.getCell(4).stringCellValue
                val agama = row.getCell(5).stringCellValue
                val pekerjaan = row.getCell(6).stringCellValue
                val kelamin = row.getCell(7).stringCellValue
                val rt = row.getCell(8).stringCellValue
                val status = row.getCell(9).stringCellValue
                val hidup = row.getCell(10).stringCellValue

                // Buat objek Penduduk dari data yang dibaca
                val penduduk = Penduduk(
                    id = 0,  // Jika id bersifat auto-increment di database, ini biasanya diabaikan.
                    nik = nik,
                    nama = nama,
                    alias = alias,
                    tempat_lahir = tempatLahir,
                    tanggal_lahir = tanggalLahir,
                    agama = agama,
                    pekerjaan = pekerjaan,
                    kelamin = kelamin,
                    rt = rt,
                    status = status,
                    hidup = hidup
                )


                // Simpan objek penduduk ke database
                lifecycleScope.launch {
                    pendudukDao.insertPenduduk(penduduk)
                }
            }

            Toast.makeText(requireContext(), "Import berhasil", Toast.LENGTH_SHORT).show()
        } catch (e: Exception) {
            Log.e("Excel Error", "Error reading Excel file: ", e)
            Toast.makeText(requireContext(), "Gagal membaca file Excel", Toast.LENGTH_SHORT).show()
        }
    }

    override fun onDestroyView() {
        super.onDestroyView()
        _binding = null
    }
}
