const express = require("express");
const dotenv = require("dotenv");
const multer = require("multer");
const cors = require("cors");
const reader = require("xlsx");
const fs = require("fs");

dotenv.config();
const app = express();
const upload = multer({ dest: "uploads/" });
const PORT = 5000;

app.use(
   cors({
      origin: [
         "http://localhost:5173",
         "https://json-to-excel-frontend-bamtb8a3e.vercel.app",
      ],
      methods: ["GET", "POST"],
      credentials: true, // Allows sending cookies
   })
);

app.use(express.json());

app.post("/upload", upload.single("file"), (req, res) => {
   const file = req.file;

   if (!file) {
      return res.status(400).send("No file uploaded.");
   }

   try {
      // Read the uploaded JSON file
      const jsonData = JSON.parse(fs.readFileSync(file.path, "utf-8"));

      // Perform the conversion logic
      const tempData = jsonData.map((company) => {
         return {
            name: company.summary.name,
            minProjectSize: company.summary.minProjectSize,
            averageHourlyRate: company.summary.averageHourlyRate,
            employees: company.summary.employees,
            address: {
               city:
                  company.summary.addresses.length > 0
                     ? company.summary.addresses[0].city
                     : "",
               state:
                  company.summary.addresses.length > 0
                     ? company.summary.addresses[0].state
                     : "",
               country:
                  company.summary.addresses.length > 0
                     ? company.summary.addresses[0].country
                     : "",
            },
            reviews: company.reviews.map((r) => {
               const solutionContent = r.content.find(
                  (c) => c.label === "SOLUTION"
               );
               // console.log('Solution Content:', solutionContent);
               return {
                  category: r.project.allCategories.join(", "),
                  description: r.project.description,
                  reviewer: {
                     title: r.reviewer.title,
                     name: r.reviewer.name,
                     location: r.reviewer.location,
                  },
                  // solution: r.content.find((c) => c.label === "SOLUTION").text
                  // solution: solutionContent ? solutionContent.text.split(" ") : []
                  solution: solutionContent.text
                     .replaceAll("?", "? \n")
                     .replaceAll("Why", "\n Why")
                     .replaceAll("What", "\n What")
                     .replaceAll("When", "\n When")
                     .replaceAll("Where", "\n Where")
                     .replaceAll("Describe", "\n Describe")
                     .replaceAll("deliverables.", "deliverables. \n"),
               };
            }),
         };
      });

      let data = [];
      tempData.forEach((company) =>
         company.reviews.forEach((review) => {
            data.push({
               name: company.name,
               category: review.category,
               description: review.description,
               reviewerTitle: review.reviewer.title,
               reviewerName: review.reviewer.name,
               reviewerLocation: review.reviewer.location,
               solution: review.solution,
            });
         })
      );

      const worksheet = reader.utils.json_to_sheet(data);
      const workbook = reader.utils.book_new();
      reader.utils.book_append_sheet(workbook, worksheet, "Sheet1");

      const outputPath = "uploads/output.xlsx";
      reader.writeFile(workbook, outputPath);

      // Send the Excel file to the user
      res.download(outputPath, "ConvertedFile.xlsx", (err) => {
         if (err) console.error(err);

         // Clean up files
         fs.unlinkSync(file.path);
         fs.unlinkSync(outputPath);
      });
   } catch (error) {
      console.error("Error processing file:", error);
      res.status(500).send("Internal server error");
   }
});

app.listen(PORT, () =>
   console.log(`Server started on http://localhost:${PORT}`)
);
