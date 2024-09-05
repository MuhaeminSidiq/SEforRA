package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	jsoniter "github.com/json-iterator/go"
	"github.com/xuri/excelize/v2"
	"html"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"
)

// Struct untuk menyimpan hasil dari API Scopus
type ScopusResponse struct {
	SearchResults struct {
		TotalResults string `json:"opensearch:totalResults"`
		StartIndex   string `json:"opensearch:startIndex"`
		ItemsPerPage string `json:"opensearch:itemsPerPage"`
		Query        struct {
			Role        string `json:"@role"`
			SearchTerms string `json:"@searchTerms"`
			StartPage   string `json:"@startPage"`
		} `json:"opensearch:Query"`
		Link []struct {
			Fa   string `json:"@_fa"`
			Ref  string `json:"@ref"`
			Href string `json:"@href"`
			Type string `json:"@type"`
		} `json:"link"`
		Entry []struct {
			Fa   string `json:"@_fa"`
			Link []struct {
				Fa   string `json:"@_fa"`
				Ref  string `json:"@ref"`
				Href string `json:"@href"`
			} `json:"link"`
			URL              string `json:"prism:url"`
			Identifier       string `json:"dc:identifier"`
			EID              string `json:"eid"`
			Title            string `json:"dc:title"`
			Creator          string `json:"dc:creator"`
			PublicationName  string `json:"prism:publicationName"`
			EISSN            string `json:"prism:eIssn"`
			Volume           string `json:"prism:volume"`
			IssueIdentifier  string `json:"prism:issueIdentifier"`
			PageRange        string `json:"prism:pageRange"`
			CoverDate        string `json:"prism:coverDate"`
			CoverDisplayDate string `json:"prism:coverDisplayDate"`
			DOI              string `json:"prism:doi"`
			CitedByCount     string `json:"citedby-count"`
			Affiliation      []struct {
				Fa                 string `json:"@_fa"`
				AffilName          string `json:"affilname"`
				AffiliationCity    string `json:"affiliation-city"`
				AffiliationCountry string `json:"affiliation-country"`
			} `json:"affiliation"`
			AggregationType    string `json:"prism:aggregationType"`
			Subtype            string `json:"subtype"`
			SubtypeDescription string `json:"subtypeDescription"`
			ArticleNumber      string `json:"article-number"`
			SourceID           string `json:"source-id"`
			OpenAccess         string `json:"openaccess"`
			OpenAccessFlag     bool   `json:"openaccessFlag"`
		} `json:"entry"`
	} `json:"search-results"`
}

// replaceXmlEscapes replaces XML escape sequences with their corresponding characters
func replaceXmlEscapes(input string) string {
	input = strings.ReplaceAll(input, `\u0026`, `&`)
	input = strings.ReplaceAll(input, `\u003C`, `<`)
	input = strings.ReplaceAll(input, `\u003E`, `>`)
	input = strings.ReplaceAll(input, `\u0022`, `"`)
	input = strings.ReplaceAll(input, `\u0027`, `'`)
	return html.UnescapeString(input)
}

// Function untuk mengambil data dari Scopus API
func fetchScopusData(doi, apiKey string) (*ScopusResponse, error) {
	apiURL := fmt.Sprintf("https://api.elsevier.com/content/search/scopus?query=DOI(%s)&apiKey=%s&httpAccept=application/json", doi, apiKey)
	resp, err := http.Get(apiURL)
	if err != nil {
		return nil, fmt.Errorf("gagal melakukan request ke API Scopus: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("gagal mengambil data dari Scopus API, status code: %d", resp.StatusCode)
	}

	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("error reading response body: %v", err)
	}

	bodyStr := string(body)
	bodyStr = replaceXmlEscapes(bodyStr)

	var scopusData ScopusResponse
	var json = jsoniter.ConfigFastest
	err = json.Unmarshal([]byte(bodyStr), &scopusData)
	if err != nil {
		return nil, fmt.Errorf("error unmarshalling JSON: %v", err)
	}

	return &scopusData, nil
}

// Function untuk parse file RIS dan ambil DOI
func parseRISFile(filePath, apiKey string) ([]ScopusResponse, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("gagal membuka file: %v", err)
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	var scopusResponses []ScopusResponse

	for scanner.Scan() {
		line := scanner.Text()
		if strings.HasPrefix(line, "DO  - ") {
			doi := strings.TrimPrefix(line, "DO  - ")
			fmt.Printf("Mengambil data untuk DOI: %s\n", doi)

			scopusData, err := fetchScopusData(doi, apiKey)
			if err != nil {
				log.Printf("Gagal mengambil data untuk DOI %s: %v\n", doi, err)
				continue
			}
			scopusResponses = append(scopusResponses, *scopusData)
		}
	}

	if err := scanner.Err(); err != nil {
		return nil, fmt.Errorf("gagal membaca file: %v", err)
	}

	return scopusResponses, nil
}

// Function untuk menyimpan hasil API ke file JSON
func saveToJSONFile(responses []ScopusResponse, outputFilePath string) error {
	data, err := json.MarshalIndent(responses, "", "  ")
	if err != nil {
		return fmt.Errorf("gagal mengonversi hasil ke JSON: %v", err)
	}

	err = ioutil.WriteFile(outputFilePath, data, 0644)
	if err != nil {
		return fmt.Errorf("gagal menulis file JSON: %v", err)
	}

	fmt.Printf("Data berhasil disimpan ke file: %s\n", outputFilePath)
	return nil
}

// Function untuk menyimpan hasil API ke file Excel
func saveToExcelFile(responses []ScopusResponse, outputFilePath string) error {
	f := excelize.NewFile()

	// Membuat sheet baru dengan nama "ScopusData"
	index, err := f.NewSheet("ScopusData")
	if err != nil {
		return fmt.Errorf("gagal membuat sheet Excel: %v", err)
	}

	// Write header
	headers := []string{"CoverDate", "DOI", "Title", "Creator", "Institution", "City", "Country", "Journal", "eISSN", "Volume", "Issue", "Pages", "OpenAccess", "CitedBy", "URL"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s1", string('A'+i))
		f.SetCellValue("ScopusData", cell, header)
	}

	// Write data
	for i, response := range responses {
		for j, entry := range response.SearchResults.Entry {
			f.SetCellValue("ScopusData", fmt.Sprintf("A%d", j+2+i), entry.CoverDate)
			f.SetCellValue("ScopusData", fmt.Sprintf("B%d", j+2+i), entry.DOI)
			f.SetCellValue("ScopusData", fmt.Sprintf("C%d", j+2+i), entry.Title)
			f.SetCellValue("ScopusData", fmt.Sprintf("D%d", j+2+i), entry.Creator)
			for _, affiliation := range entry.Affiliation {
				f.SetCellValue("ScopusData", fmt.Sprintf("E%d", j+2+i), affiliation.AffilName)
				f.SetCellValue("ScopusData", fmt.Sprintf("F%d", j+2+i), affiliation.AffiliationCity)
				f.SetCellValue("ScopusData", fmt.Sprintf("G%d", j+2+i), affiliation.AffiliationCountry)
			}
			f.SetCellValue("ScopusData", fmt.Sprintf("H%d", j+2+i), entry.PublicationName)
			f.SetCellValue("ScopusData", fmt.Sprintf("I%d", j+2+i), entry.EISSN)
			f.SetCellValue("ScopusData", fmt.Sprintf("J%d", j+2+i), entry.Volume)
			f.SetCellValue("ScopusData", fmt.Sprintf("K%d", j+2+i), entry.IssueIdentifier)
			f.SetCellValue("ScopusData", fmt.Sprintf("L%d", j+2+i), entry.PageRange)
			f.SetCellValue("ScopusData", fmt.Sprintf("M%d", j+2+i), entry.OpenAccess)
			f.SetCellValue("ScopusData", fmt.Sprintf("N%d", j+2+i), entry.CitedByCount)
			f.SetCellValue("ScopusData", fmt.Sprintf("O%d", j+2+i), entry.URL)
		}
	}

	// Set active sheet ke "ScopusData"
	f.SetActiveSheet(index)

	// Simpan file Excel
	err = f.SaveAs(outputFilePath)
	if err != nil {
		return fmt.Errorf("gagal menyimpan file Excel: %v", err)
	}

	fmt.Printf("Data berhasil disimpan ke file: %s\n", outputFilePath)
	return nil
}

func main() {
	// Meminta input path file RIS dan API key dari pengguna
	fmt.Print("Masukkan path file RIS: ")
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
	filePath := scanner.Text()

	fmt.Print("Masukkan API key Scopus: ")
	scanner.Scan()
	apiKey := scanner.Text()

	if strings.TrimSpace(filePath) == "" || strings.TrimSpace(apiKey) == "" {
		log.Fatal("Path file RIS dan API key Scopus tidak boleh kosong.")
	}

	// Parse file RIS dan ambil data dari Scopus
	scopusResponses, err := parseRISFile(filePath, apiKey)
	if err != nil {
		log.Fatalf("Gagal memproses file RIS: %v", err)
	}

	// Tentukan nama file output berdasarkan nama file RIS
	jsonOutputFilePath := filepath.Join(filepath.Dir(filePath), strings.TrimSuffix(filepath.Base(filePath), filepath.Ext(filePath))+".json")
	excelOutputFilePath := filepath.Join(filepath.Dir(filePath), strings.TrimSuffix(filepath.Base(filePath), filepath.Ext(filePath))+".xlsx")

	// Simpan hasil ke file JSON
	err = saveToJSONFile(scopusResponses, jsonOutputFilePath)
	if err != nil {
		log.Fatalf("Gagal menyimpan hasil ke file JSON: %v", err)
	}

	// Simpan hasil ke file Excel
	err = saveToExcelFile(scopusResponses, excelOutputFilePath)
	if err != nil {
		log.Fatalf("Gagal menyimpan hasil ke file Excel: %v", err)
	}
}
