using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.IO;
using ExcelDataReader;
using System.Data;
using UnityEngine.Networking;

public class ExcelReader : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        StartCoroutine(ReadExcel());
    }

    IEnumerator ReadExcel()
    {
        // Path to your Excel file
        
        string path = "C:/Users/Win10/Desktop/deneme.xlsx"; // bunu doðru formatta yazdýðýndan emin ol en ufak hatada çalýþmýyo :(

        // Download the file
        UnityWebRequest www = UnityWebRequest.Get(path);
        yield return www.SendWebRequest();

        if (www.result != UnityWebRequest.Result.Success)
        {
            Debug.LogError(www.error);
            yield break;
        }

        // Read the Excel data
        byte[] excelData = www.downloadHandler.data;

        using (MemoryStream stream = new MemoryStream(excelData))
        {
            using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                // Choose the sheet index (0 for the first sheet),
                while (reader.Read())
                {
                    // Assuming the first row is the header
                    if (reader.Depth == 0)
                    {
                        // Process header (column names)
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            Debug.Log(reader.GetString(i));
                        }
                    }
                    else
                    {
                        // Process data rows
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            Debug.Log(reader.GetValue(i));
                        }
                    }
                }
            }
        }
    }
}

