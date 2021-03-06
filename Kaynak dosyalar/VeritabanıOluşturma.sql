USE [master]
GO
/****** Object:  Database [Sinav_Olusturma]    Script Date: 5.06.2017 14:45:40 ******/
CREATE DATABASE [Sinav_Olusturma]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'Sinav_Olusturma', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL11.SQLEXPRESS\MSSQL\DATA\Sinav_Olusturma.mdf' , SIZE = 4160KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'Sinav_Olusturma_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL11.SQLEXPRESS\MSSQL\DATA\Sinav_Olusturma_log.ldf' , SIZE = 1040KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO
ALTER DATABASE [Sinav_Olusturma] SET COMPATIBILITY_LEVEL = 100
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [Sinav_Olusturma].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [Sinav_Olusturma] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET ARITHABORT OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET AUTO_CREATE_STATISTICS ON 
GO
ALTER DATABASE [Sinav_Olusturma] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [Sinav_Olusturma] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [Sinav_Olusturma] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET  DISABLE_BROKER 
GO
ALTER DATABASE [Sinav_Olusturma] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [Sinav_Olusturma] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET RECOVERY FULL 
GO
ALTER DATABASE [Sinav_Olusturma] SET  MULTI_USER 
GO
ALTER DATABASE [Sinav_Olusturma] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [Sinav_Olusturma] SET DB_CHAINING OFF 
GO
ALTER DATABASE [Sinav_Olusturma] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [Sinav_Olusturma] SET TARGET_RECOVERY_TIME = 0 SECONDS 
GO
USE [Sinav_Olusturma]
GO
/****** Object:  Table [dbo].[Dersler]    Script Date: 5.06.2017 14:45:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Dersler](
	[ders_id] [int] IDENTITY(1,1) NOT NULL,
	[ders_adi] [nchar](100) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Kullanicilar]    Script Date: 5.06.2017 14:45:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Kullanicilar](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[kullanici] [nchar](10) NULL,
	[sifre] [nchar](10) NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[Sorular]    Script Date: 5.06.2017 14:45:40 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Sorular](
	[soru_id] [int] IDENTITY(1,1) NOT NULL,
	[ders_id] [int] NULL,
	[soru] [nvarchar](max) NULL,
	[soru_resim] [nvarchar](max) NULL,
	[a_cevap] [nvarchar](max) NULL,
	[b_cevap] [nvarchar](max) NULL,
	[c_cevap] [nvarchar](max) NULL,
	[d_cevap] [nvarchar](max) NULL,
	[e_cevap] [nvarchar](max) NULL,
	[dogru_cevap] [nvarchar](max) NULL,
	[zorluk_derecesi] [nchar](10) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
USE [master]
GO
ALTER DATABASE [Sinav_Olusturma] SET  READ_WRITE 
GO
