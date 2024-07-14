const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');
const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.post('/submit', (req, res) => {
    const data = [
        {
            Customer: req.body.customer,
            CrewLead: req.body.crewLead,
            PO: req.body.po,
            ReviewWorkOrder: req.body.reviewWorkOrder,
            SpecialAreas: req.body.specialAreas,
            ProjectExpectations: req.body.projectExpectations,
            EstimatedTimelineStart: req.body.estimatedTimelineStart,
            EstimatedTimelineEnd: req.body.estimatedTimelineEnd,
            WorkingHoursStart: req.body.workingHoursStart,
            WorkingHoursEnd: req.body.workingHoursEnd,
            AccessDetails: req.body.accessDetails,
            PMCommunication: req.body.pmCommunication,
            PMContact: req.body.pmContact,
            CustomerContact: req.body.customerContact,
            InspectionPayment: req.body.inspectionPayment,
            Signature: req.body.signature,
            Date: req.body.date,
            SidingProduct: req.body.productSiding,
            SidingColor: req.body.colorSiding,
            SidingSheen: req.body.sheenSiding,
            // Add more fields as needed
        }
    ];

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
    XLSX.writeFile(workbook, 'form-data.xlsx');

    res.status(200).send('Form data submitted and saved to Excel file successfully!');
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}/`);
});
