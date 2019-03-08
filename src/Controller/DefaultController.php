<?php

namespace App\Controller;

use Symfony\Component\Routing\Annotation\Route;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Doctrine\DBAL\Driver\Connection;
use Unirest\Request as RestRequest;
use Unirest\Response as RestResponse;
use Unirest\Request\Body;
use Symfony\Component\DependencyInjection\ParameterBag\ParameterBagInterface;
use App\Entity\User;
use Symfony\Component\Form\Extension\Core\Type\TextType;
use Symfony\Component\Form\Extension\Core\Type\EmailType;
use Symfony\Component\Form\Extension\Core\Type\SubmitType;
use Symfony\Component\Form\Extension\Core\Type\CheckboxType;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class DefaultController extends AbstractController
{
    

    /**
     * @Route("/", name="homepage")
     */
    public function index()
    {
        return $this->render('default/homepage.html.twig');
    }


    /**
     * @Route("/form", name="form")
     */
    public function form(Request $request)
    {
        $user = new User();


        $form = $this->createFormBuilder($user)
            ->add('nom', TextType::class)
            ->add('prenom', TextType::class)
            ->add('email', EmailType::class)
            ->add('newsletter', CheckboxType::class)
            ->add('save', SubmitType::class, ['label' => 'l'])
            ->getForm();


        $form->handleRequest($request);

        if ($form->isSubmitted() && $form->isValid()) {
            // $form->getData() holds the submitted values
            // but, the original `$task` variable has also been updated
            $user = $form->getData();

            // ... perform some action, such as saving the task to the database
            // for example, if Task is a Doctrine entity, save it!
             $em = $this->getDoctrine()->getManager();
            $em->persist($user);
            $em->flush();

            $this->addFlash('success', 'Bonne Chance ! Votre participation a bien été prise en compte');
            return $this->redirectToRoute('homepage');
        }

        return $this->render('default/form.html.twig', [
            'form' => $form->createView(),
        ]);
    }

    /**
     * @Route("/excel", name="excel")
     */
    public function excel()
    {
        return $this->render('default/excel.html.twig');
    }

    /**
     * @Route("/extraxt", name="extraxt")
     */
    public function extraxt()
    {
        $spreadsheet = new Spreadsheet();

        /* @var $sheet \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Worksheet */
        $em = $this->getDoctrine()->getManager();
        $users = $em->getRepository('App:User')->findAll();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Nom');
        $sheet->setCellValue('B1', 'Prénom');
        $sheet->setCellValue('C1', 'Email');
        $sheet->setCellValue('D1', 'Newsletter');
        foreach($users as $k=>$user)
        {
            $sheet->setCellValue('A'.strval($k+2), $user->getNom());
            $sheet->setCellValue('B'.strval($k+2), $user->getPrenom());
            $sheet->setCellValue('C'.strval($k+2), $user->getEmail());
            if($user->getNewsletter())
            {
                $sheet->setCellValue('D'.strval($k+2),'Oui');
            }
            else{
                $sheet->setCellValue('D'.strval($k+2),'Non');
            }

        }


        $sheet->setTitle("Jeu EAFIT");

        // Create your Office 2007 Excel (XLSX Format)
        $writer = new Xlsx($spreadsheet);

        // Create a Temporary file in the system
        $fileName = 'jeu_eafit.xlsx';
        $temp_file = tempnam(sys_get_temp_dir(), $fileName);

        // Create the excel file in the tmp directory of the system
        $writer->save($temp_file);

        // Return the excel file as an attachment
        return $this->file($temp_file, $fileName, ResponseHeaderBag::DISPOSITION_INLINE);
    }
}
