using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace BreakpointSetter
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<string> LProyectsName = new List<string>();
                LProyectsName.Add("xxx");
                LProyectsName.Add("xxx");
                LProyectsName.Add("xxx");

                foreach (string projectName in LProyectsName)
                {
                    // Obtener una referencia a la instancia actual de Visual Studio
                    DTE2 dte = GetRunningInstance();
                    if (dte == null)
                    {
                        Console.WriteLine("No se encontró ninguna instancia de Visual Studio.");
                        return;
                    }

                    // Obtener la solución actual
                    Solution solution = dte.Solution;

                    // Buscar el proyecto por nombre exacto
                    Project targetProject = FindProjectRecursively(solution.Projects, projectName);

                    if (targetProject == null)
                    {
                        Console.WriteLine($"El proyecto con el nombre '{projectName}' no fue encontrado.");
                        return;
                    }

                    Console.WriteLine($"Proyecto encontrado: {targetProject.Name}");

                    // Recorrer todos los archivos del proyecto y establecer puntos de interrupción
                    foreach (ProjectItem projectItem in targetProject.ProjectItems)
                    {
                        ProcessProjectItem(projectItem);
                    }

                    Console.WriteLine("Puntos de interrupción añadidos a todos los métodos.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Mantener la ventana abierta
                Console.WriteLine("Presiona cualquier tecla para salir...");
                Console.ReadLine();
            }
        }

        static DTE2 GetRunningInstance()
        {
            try
            {
                CreateBindCtx(0, out var ctx);
                ctx.GetRunningObjectTable(out var rot);
                rot.EnumRunning(out var enumMoniker);
                enumMoniker.Reset();

                var monikers = new IMoniker[1];
                while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    rot.GetObject(monikers[0], out var comObject);
                    var dte = comObject as DTE2;
                    if (dte != null)
                    {
                        return dte;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return null;
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        static Project FindProjectRecursively(Projects projects, string projectName)
        {
            try
            {
                foreach (Project project in projects)
                {
                    if (project.Name.Equals(projectName, StringComparison.OrdinalIgnoreCase))
                    {
                        return project;
                    }

                    // Si el proyecto tiene subitems (carpetas), recorrer recursivamente
                    if (project.ProjectItems != null)
                    {
                        Project foundProject = FindProjectInProjectItems(project.ProjectItems, projectName);
                        if (foundProject != null)
                        {
                            return foundProject;
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

        static Project FindProjectInProjectItems(ProjectItems projectItems, string projectName)
        {
            try
            {
                foreach (ProjectItem item in projectItems)
                {
                    // Si es una carpeta o un proyecto
                    if (item.SubProject != null)
                    {
                        if (item.SubProject.Name.Equals(projectName, StringComparison.OrdinalIgnoreCase))
                        {
                            return item.SubProject;
                        }

                        // Llamar recursivamente si hay más subcarpetas o proyectos
                        Project foundProject = FindProjectInProjectItems(item.SubProject.ProjectItems, projectName);
                        if (foundProject != null)
                        {
                            return foundProject;
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

        static void ProcessProjectItem(ProjectItem item)
        {
            try
            {
                // Si el archivo tiene un modelo de código, procesar sus elementos
                if (item.FileCodeModel != null)
                {
                    foreach (CodeElement element in item.FileCodeModel.CodeElements)
                    {
                        if (element.Kind == vsCMElement.vsCMElementNamespace)
                        {
                            CodeNamespace ns = (CodeNamespace)element;
                            ProcessNamespace(ns);
                        }
                    }
                }

                // Recursivamente procesar los subelementos
                if (item.ProjectItems != null)
                {
                    foreach (ProjectItem subItem in item.ProjectItems)
                    {
                        ProcessProjectItem(subItem);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void ProcessNamespace(CodeNamespace ns)
        {
            int retryCount = 5; // Número de reintentos permitidos
            while (retryCount > 0)
            {
                try
                {
                    foreach (CodeElement element in ns.Members)
                    {
                        if (element.Kind == vsCMElement.vsCMElementClass)
                        {
                            CodeClass cls = (CodeClass)element;
                            ProcessClass(cls);
                        }
                    }
                    break; // Si el proceso fue exitoso, salir del bucle
                }
                catch (COMException ex) when ((uint)ex.ErrorCode == 0x8001010A) // Código específico de RPC_E_SERVERCALL_RETRYLATER
                {
                    Console.WriteLine("Visual Studio está ocupado. Reintentando en 500ms...");
                    System.Threading.Thread.Sleep(500); // Esperar medio segundo antes de reintentar
                    retryCount--;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    break;
                }
            }

            if (retryCount == 0)
            {
                Console.WriteLine("Se agotaron los reintentos. Visual Studio sigue ocupado.");
            }
        }

        static void ProcessClass(CodeClass cls)
        {
            try
            {
                foreach (CodeElement element in cls.Members)
                {
                    if (element.Kind == vsCMElement.vsCMElementFunction)
                    {
                        CodeFunction func = (CodeFunction)element;
                        AddBreakpoint(func);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void AddBreakpoint(CodeFunction func)
        {
            try
            {
                TextPoint startPoint = func.GetStartPoint(vsCMPart.vsCMPartBody);
                func.DTE.Debugger.Breakpoints.Add("", func.ProjectItem.FileNames[1], startPoint.Line, 1, "",
                    dbgBreakpointConditionType.dbgBreakpointConditionTypeWhenTrue, "", "", 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
